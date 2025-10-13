
# Columns in the `Data Entry` sheets

| Column name | Description | Changeable | Valid values | Provenance|
|-------------|-------------|------------|--------------|-----------|
|material_id|The unique id for this item defined in CRC (copyrightcheck)|No|-|CRC|
|period|When the course was given, e.g. `2025-1A`|No|-|CRC/Canvas|
|is_duplicate|Whether this file was found more than once. Currently not implemented.|No|-|Custom script|
|url|The URL where the item can be accessed in Canvas|No|-|CRC/Canvas|
|file_exists|Whether the file still exists on Canvas at the time of the last check.|No|Yes, No|Custom script/Canvas|
|latest_filecheck_date|The date when the file was last checked|No|-|Custom script|
|workflow_status|The current status of the workflow for this item|Yes|ToDo, InProgress, Done|Custom script|
|v1_manual_classification|Manual classification using v1 style|Yes|open access, korte overname, middellange overname, lange overname, eigen materiaal - powerpoint, eigen materiaal - titelindicatie, eigen materiaal - overig, eigen materiaal, onbekend, niet geanalyseerd, in onderzoek, verwijderverzoek verstuurd, licentie beschikbaar|user entry|
|v2_manual_classification|Manual classification using v2 style|Yes|Ja (open licentie), Ja (bibilotheek licentie), Ja (directe toestemming), Ja (Publiek domein), Ja (eigen werk), Ja (studentwerk), Ja (easy access), Ja (anders), Ja (directe toestemming) - tijdelijk, Ja (bibilotheek licentie)- tijdelijk, Ja (anders) - tijdelijk, Nee (Link beschikbaar), Nee (studentwerk), Nee, Onbekend|user entry/converted using script|
|v2_length|The length of the item (v2)|No|Kort, Middellang, Lang, Onbekend|user entry/converted using script|
|v2_overname_status|The overname status (v2)|No|Overname (inbreukmakende), Overname (andere), Geen overname, Onbekend|user entry/converted using script|
|remarks|Any remarks made by the user|Yes|any text|user entry|
|ml_prediction|the classification predicted by CRC using ML model|No|see v1 classification|CRC|
|pdf_filename|the filename of this pdf|No|-|CRC|
|pdf_title|the title of the document extracted from the pdf|No|-|CRC|
|uploaded_by|the user/teacher who uploaded the document|No|-|CRC/Canvas|
|detected_pdf_author|the name of the author detected in the pdf|No|-|CRC|
|course_contacts_names|the name(s) of the course contact(s)|No|-|OSIRIS/Custom script|
|course_contacts_emails|the email(s) of the course contact(s)|No|-|OSIRIS/Custom script|
|course_name_canvas|the name of the course in Canvas|No|-|CRC|
|course_names_osiris|the name of the course in OSIRIS|No|-|OSIRIS/Custom script|
|programme_canvas|the name of the related programme in Canvas|No|-|CRC|
|programmes_osiris|the name of the related programme in OSIRIS|No|-|OSIRIS/Custom script|
