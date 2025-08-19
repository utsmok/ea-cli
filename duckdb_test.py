import marimo

__generated_with = "0.14.17"
app = marimo.App(width="full")


@app.cell
def _():
    import duckdb
    import marimo as mo
    DATABASE_URL = "testing.duckdb"
    engine = duckdb.connect(DATABASE_URL, read_only=False)
    return engine, mo


@app.cell
def _(copyright_data, engine, mo, pdf_data, text_embeddings):
    _df = mo.sql(
        f"""
        COPY (
          SELECT
            cd.*,
            pd.*,
            te.*
          FROM copyright_data AS cd
          LEFT JOIN pdf_data AS pd
            ON cd.material_id = pd.material_id
          LEFT JOIN text_embeddings as te
            ON cd.material_id = te.material_id
        )    
            TO 'copyright_data_with_pdf_embeddings.parquet' (FORMAT 'parquet');

        """,
        engine=engine
    )
    return


@app.cell
def _(engine, mo, pdf_data):
    _df = mo.sql(
        f"""
        select count(material_id) from pdf_data
        """,
        engine=engine
    )
    return


@app.cell
def _():
    from sentence_transformers import SentenceTransformer

    model = SentenceTransformer('all-MiniLM-L6-v2')

    def get_text_embedding_list(list_text: list[str]):
        """
        Return the list of normalized vector embeddings for list_text.
        """
        return model.encode(list_text, normalize_embeddings=True)


    return (get_text_embedding_list,)


@app.cell
def _(engine, get_text_embedding_list):
    engine.create_function(
        "get_text_embedding_list",
        get_text_embedding_list,
        return_type='FLOAT[384][]'
    )

    return


@app.cell
def _(engine, mo):
    _df = mo.sql(
        f"""
        -- split the text_embedding array

        """,
        engine=engine
    )
    return


@app.cell
def _(engine):
    engine.sql("""
        create table text_embeddings (
            material_id integer,
            text_embedding FLOAT[384]
        )
    """)
    return


@app.cell
def _(engine, mo):
    _df = mo.sql(
        f"""

        """,
        engine=engine
    )
    return


@app.cell
def _(engine):
    total=2131
    num_batches = 250
    batch_size = total // num_batches + 1
    for i in range(num_batches):
        selection_query = (
            engine.table("pdf_data")
            .order("material_id")
            .limit(batch_size, offset=batch_size*i)
            .select("*")
        )

        (
            selection_query.aggregate("""
                array_agg(extracted_text) as text_list,
                array_agg(material_id) as id_list,
                get_text_embedding_list(text_list) as text_emb_list
            """).select("""
                unnest(id_list) as material_id,
                unnest(text_emb_list) as text_embedding
            """)
        ).insert_into("text_embeddings")


    return


if __name__ == "__main__":
    app.run()
