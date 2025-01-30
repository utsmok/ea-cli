import polars as pl
from constants import COURSE_MAPPING, DEPARTMENT_MAPPING, FINE_AMOUNT, DIRS
from enrichment import enrich_df_with_osiris_data
from utils import info, warn, cool, Directory, File
from datetime import datetime
import locale
from sheet import finalize_sheet
from rich.console import Console
from rich.table import Table
from rich.terminal_theme import SVG_EXPORT_THEME
import copy

def create_programme_overviews(all_faculty_data: pl.DataFrame, faculty: str, style_iter:int):
    """
    create an overview sheet for each programme of the given faculty, using the data in df.
    """
    course_to_group: dict[str,str] = COURSE_MAPPING[faculty]
    data: list[dict[str,pl.DataFrame]] = []
    today = datetime.now().strftime("%Y-%m-%d")

    for course, group in course_to_group.items():
        programme_data = all_faculty_data.filter(pl.col("department") == course)
        if programme_data.is_empty():
            continue
        else:
            programme_data = programme_data.with_columns(
                pl.col('pages_x_students').cast(pl.Int32).mul(FINE_AMOUNT).alias('possible_fine')
            )
            programme_data = programme_data.with_columns(
                infringement=pl.when(pl.col("manual_classification").is_null() |
                                    (pl.col("manual_classification") == "") |
                                    (pl.col("manual_classification") == "-"))
                                .then(pl.lit("undetermined"))
                                .when(pl.col("manual_classification").str.to_lowercase().str.contains("open|eigen|overig|deleted"))
                                .then(pl.lit("no"))
                                .when(pl.col("manual_classification").str.to_lowercase().str.contains("lange"))
                                .then(pl.lit("yes"))
                                .otherwise(pl.lit("maybe"))
            )
            data.append({'group':group, 'data': programme_data})

    final_data: dict[str,pl.DataFrame] = {}
    for item in data:
        info(f'group: {item.get("group")} --> + {item.get("data").shape[0]} items')
        if item.get('group') in final_data:
            final_data[item.get('group')] = pl.concat([final_data[item.get('group')], item['data']])
        else:
            final_data[item.get('group')] = item['data']

    overview_fac_programme_dir = Directory(DIRS['overviews_backup'].full / faculty / "per_programme")
    for groupname, df in final_data.items():
        for file in Directory(DIRS['faculties'].full / faculty / "per_programme").files:
            if file.extension not in [".xls", ".xlsx"]:
                continue
            if 'overview' in file.name and groupname in file.name:
                file.move( overview_fac_programme_dir.full / file.name)
                continue
        info(f'{groupname} has {df.shape[0]} items')
        programme_file = File(DIRS['faculties'].full / faculty / "per_programme" / f'{groupname}_total_overview_updated_{today}.xlsx')
        info(f'saving file with {df.shape[0]} rows to {programme_file.path}')
        programme_data.write_excel(programme_file.path)
        style_iter = finalize_sheet(programme_file, programme_data, style_iter)

    return style_iter



def create_faculty_overviews(faculty_data: dict[str, pl.DataFrame], style_iter:int) -> int:
    """
    Input:
    faculty_data: dict[str, pl.DataFrame]:
        keys are the faculty names (BMS, EEMCS, etc)
        values are the dataframes with the data for each faculty
    """
    # loop over the faculties
    # for each, read in all data and store
    overview_data: list[dict] = []
    today = datetime.now().strftime("%Y-%m-%d")

    for faculty, all_faculty_data in faculty_data.items():
        all_faculty_data = enrich_df_with_osiris_data(all_faculty_data, faculty)
        if faculty in COURSE_MAPPING:
            style_iter = create_programme_overviews(all_faculty_data, faculty, style_iter)
        fac_data = {'faculty': faculty}

        # add columns:
        # 'possible_fine': for each row multiply col pages_x_students with 0.30 to get the amount
        if all_faculty_data.is_empty():
            continue

        all_faculty_data = all_faculty_data.with_columns(
            pl.col('pages_x_students').cast(pl.Int32).mul(FINE_AMOUNT).alias('possible_fine')
        )

        # 'infringement': possible values: 'yes', 'no', 'maybe', 'undetermined'.
        # based on the value in 'manual_classification'
        # if 'manual_classification' is empty (None, "", '-', NaN): set to 'undetermined'
        # if the str in 'manual_classification' contains 'open' or 'eigen': set no 'no'
        # if 'lange overname' is in 'manual_classification': set 'yes'
        # else set to 'maybe'

        all_faculty_data = all_faculty_data.with_columns(
            infringement=pl.when(pl.col("manual_classification").is_null() |
                                (pl.col("manual_classification") == "") |
                                (pl.col("manual_classification") == "-"))
                            .then(pl.lit("undetermined"))
                            .when(pl.col("manual_classification").str.to_lowercase().str.contains("open|eigen|overig|deleted"))
                            .then(pl.lit("no"))
                            .when(pl.col("manual_classification").str.to_lowercase().str.contains("lange"))
                            .then(pl.lit("yes"))
                            .otherwise(pl.lit("maybe"))
        )

        # calculate the total possible fine by adding up all values in the 'possible_fine' column
        # for all items that do not have 'no' in the 'infringement' column

        total_possible_fine = all_faculty_data.filter(pl.col("infringement") != "no").select(pl.sum('possible_fine')).to_series().to_list()[0]
        definitive_fine = all_faculty_data.filter(pl.col("infringement") == "yes").select(pl.sum('possible_fine')).to_series().to_list()[0]
        locale.setlocale(locale.LC_ALL, 'nl_NL.utf8')
        fac_data['total_possible_fine'] = str(locale.currency(total_possible_fine, grouping=True, symbol=True))
        fac_data['definitive_fine']= str(locale.currency(definitive_fine, grouping=True, symbol=True))
        fac_data['items_total'] = str(all_faculty_data.shape[0])
        fac_data['possible_infringements'] = str(all_faculty_data.filter(pl.col("infringement") != "no").shape[0])
        fac_data['definitive_infringements'] = str(all_faculty_data.filter(pl.col("infringement") == "yes").shape[0])
        fac_data['definitive_non_infringements'] = str(all_faculty_data.filter(pl.col("infringement") == "no").shape[0])
        fac_data['items_without_man_cl'] = str(all_faculty_data.filter(pl.col("infringement") == "undetermined").shape[0])
        fac_data['items_to_do'] = str(all_faculty_data.filter(pl.col("workflow_status") == "ToDo").shape[0])
        overview_data.append(fac_data)
        fac_file = File(DIRS['faculties'].full / faculty / f'{faculty}_total_overview_updated_{today}.xlsx')
        info(f'saving file with {all_faculty_data.shape[0]} rows to {fac_file.path}')
        all_faculty_data.write_excel(fac_file.path)
        style_iter = finalize_sheet(fac_file, all_faculty_data, style_iter)
        locale.setlocale(locale.LC_ALL, '')

    # now we have the data for all faculties, and written the excel files to disk.
    # print the overview table to the console, and export it as an html file to the faculties/overviews dir.
    cons = Console(record=True)

    datatable = Table(title=f'Faculty Overview {today}')
    datatable.add_column('Faculty', justify='right', style='yellow bold')
    datatable.add_column('Probable fine', justify='left', style='red bold')
    datatable.add_column('Max fine', justify='left')
    datatable.add_column('Items total', justify='center', style='cyan bold')
    datatable.add_column('Infringements', justify='center')
    datatable.add_column('Non-infringements', justify='center')
    datatable.add_column('To be classified', justify='center', style='magenta bold')
    datatable.add_column('To do', justify='center', style='magenta bold')
    factable = copy.deepcopy(datatable)
    for fac in overview_data:
        # save html overview for each faculty in their dir
        # also add that data to the overview html
        cur_fac_table = copy.deepcopy(factable)
        cur_fac_table.add_row(fac['faculty'],
                        fac['definitive_fine'],
                        fac['total_possible_fine'],
                        fac['items_total'],
                        fac['definitive_infringements']+f" ({int(fac['definitive_infringements'])/int(fac['items_total'])*100:.0f}%)",
                        fac['definitive_non_infringements']+f" ({int(fac['definitive_non_infringements'])/int(fac['items_total'])*100:.0f}%)",
                        fac['items_without_man_cl']+f" ({int(fac['items_without_man_cl'])/int(fac['items_total'])*100:.0f}%)",
                        fac['items_to_do'] + f" ({int(fac['items_to_do'])/int(fac['items_total'])*100:.0f}%)"
        )
        cons.print(cur_fac_table)
        cons.print('''Explanation of columns:

            - [yellow bold]Faculty[/yellow bold]: the abbreviation of the faculty -- all data is per faculty
            - [red bold]Probable fine[/red bold]: the sum of all fines for items that are manually classified as 'lange overname'
            - [bold]Max fine[/bold]: the sum of all fines for all items except those manually classified as 'eigen materiaal' or 'open access'
            - [cyan bold]Items total[/cyan bold]: the total number of items selected by the 'CopyRight tool' (i.e. all pdfs with 40+ pages)
            - [bold]Infringements[/bold]: the number of items that are manually classified as 'lange overname' -- plus as a percentage of total number of items
            - [bold]Non-infringements[/bold]: the number of items manually classified as 'eigen materiaal' or 'open access' -- plus as a percentage of total number of items
            - [magenta bold]To be classified[/magenta bold]: the number of items that are not yet manually classified -- plus as a percentage of total number of items
            ''')
        facdir = Directory(DIRS['faculties'].full / fac['faculty'])
        # delete any old html files

        for file in facdir.files:
            if file.name.endswith('.html'):
                file.delete()

        cons.save_html(facdir.full / f'summary_{today}.html', theme=SVG_EXPORT_THEME)

        datatable.add_row(fac['faculty'],
                        fac['definitive_fine'],
                        fac['total_possible_fine'],
                        fac['items_total'],
                        fac['definitive_infringements']+f" ({int(fac['definitive_infringements'])/int(fac['items_total'])*100:.0f}%)",
                        fac['definitive_non_infringements']+f" ({int(fac['definitive_non_infringements'])/int(fac['items_total'])*100:.0f}%)",
                        fac['items_without_man_cl']+f" ({int(fac['items_without_man_cl'])/int(fac['items_total'])*100:.0f}%)",
                        fac['items_to_do'] + f" ({int(fac['items_to_do'])/int(fac['items_total'])*100:.0f}%)"

                    )

    # now save the complete table to all_items

    cons.print(datatable)
    cons.print('''Explanation of columns:

            - [yellow bold]Faculty[/yellow bold]: Faculty abbreviation
            - [red bold]Probable fine[/red bold]: Total fine for items that have 'lange overname' as manual classification
            - [bold]Max fine[/bold]: Total fine for all items excluding items manually classified as 'eigen materiaal' or 'open access'
            - [cyan bold]Items total[/cyan bold]: Total amount of 'lange overnames' found by the 'CopyRight tool' (all pdfs with 40+ pages)
            - [bold]Infringements[/bold]: Items manually classified as 'lange overname', (% of total)
            - [bold]Non-infringements[/bold]: Items manually classified as 'eigen materiaal' or 'open access', (% of total)
            - [magenta bold]To be classified[/magenta bold]: Items not yet manually classified, (% of total)
            - [magenta bold]To do[/magenta bold]: Items in need of action by faculty, (% of total)
            ''')
    cons.save_html(DIRS['all_items'].full / f'faculty_overview_{today}.html', theme=SVG_EXPORT_THEME)
    return style_iter
