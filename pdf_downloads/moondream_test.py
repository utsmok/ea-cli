import marimo

__generated_with = "0.13.11"
app = marimo.App(width="full")


@app.cell
def _():
    import marimo as mo
    import moondream as md
    from PIL import Image
    from pathlib import Path
    import pymupdf
    import timeit
    return Image, Path, md, pymupdf


@app.cell
def _(md):
    model = md.vl(
        api_key="eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJrZXlfaWQiOiJiYzlkMzFiZS1jNTk5LTQyNzMtOWJmMC1jZjIyMTE3OGI2ZmEiLCJvcmdfaWQiOiJnMXBPMVcxMjJzbWRaSU5ZUGFjUUZicURiVWpON2d0VCIsImlhdCI6MTc0NzkxNDk3NiwidmVyIjoxfQ.wIm54chnOX_38P-Am7Xar8OWP0xTBclB9SzbjSRzGo4"
    )
    return (model,)


@app.cell
def _(Image, Path, pymupdf):
    def get_pdfs_as_img(dir:str = '.', max_items:int = 50, imgs_per_doc: int = 1) -> dict[str, list[Image]]:
        pdf_files = Path(dir).glob('*.pdf')
        images:dict[str,dict[str, list[Image]|int]] = {}
        skip_amount= 50
        for docnum, file in enumerate(pdf_files):
            skip_amount -= 1
            if skip_amount > 0:
                continue
            filename = file.name
            material_id, filename = filename.split('_', maxsplit=1)
            filename = filename.split('_', maxsplit=1)[1].split('.pdf_')[0]
            for i in range(1,10):
                filename=filename.replace(f'-{i}','')
            filename = filename.strip().replace(' ', '_')
            if filename in images:
                continue
            try:
                doc = pymupdf.open(file)
            except Exception as e:
                print(f'failed to open {file}: {e}.\n skipping.')
                continue
            images[filename]={'id':int(material_id), 'images':[]}
            for page in doc:           
                images[filename]['images'].append(page.get_pixmap().pil_image())
                if len(images[filename]['images']) >= imgs_per_doc:
                    break
            if len(images) >= max_items:
                break
    
        print(f'got {len(images)} 1st page images from pdfs')
        return images
    return (get_pdfs_as_img,)


@app.cell
def _(get_pdfs_as_img, model):
    def enrich_with_moondream():
        pdfs_as_imgs = get_pdfs_as_img()
        parsed = {}
    
        for filename, data in pdfs_as_imgs.items():
            if len(data['images']) == 0:
                continue
            data['captions'] = []
            for img_num, img in enumerate(data['images']):
                try:
                    data['captions'].append(model.caption(img, length="short")["caption"])
                    print(data['captions'][-1])
                except Exception:
                    print(f'failed to get caption, skipping')
                if img_num==0:
                    try:
                        data['type']=model.query(img, "what type of document is this?")
                        print(data['type'])
                    except Exception:
                        print(f'failed to get type, skipping')
                    try:
                        data['authors']=model.query(img, "who are the authors of this document?")
                        print(data['authors'])
                    except Exception:
                            print(f'failed to get authors, skipping')
                    #try:
                        #data['transcribed']=model.query(img, "Transcribe the text in natural reading order")
                        #print(data['transcribed'])
                    #except Exception:
                        #print(f'failed to get transcription, skipping')
                    
            parsed[filename] = data
    
            if len(parsed) >= 5:
                break
        return parsed


    enriched = enrich_with_moondream()
        
        
    return (enriched,)


@app.cell
def _(enriched):
    enriched
    return


if __name__ == "__main__":
    app.run()
