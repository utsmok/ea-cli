CREATE TABLE IF NOT EXISTS "item_updates" (
    "created_at" TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    "modified_at" TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    "id" INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
    "change_details" JSON NOT NULL,
    "material_id" INT NOT NULL
);
CREATE TABLE sqlite_sequence(name,seq);
CREATE TABLE IF NOT EXISTS "llm_classification_data" (
    "created_at" TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    "modified_at" TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    "id" INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
    "allowed_usage" VARCHAR(255) NOT NULL DEFAULT 'undetermined' /* ALLOWED: allowed\nRESTRICTED: restricted\nNOT_ALLOWED: not allowed\nUNDETERMINED: undetermined */,
    "allowed_usage_reasoning" VARCHAR(10000) NOT NULL,
    "copyright_status" VARCHAR(255) NOT NULL DEFAULT 'other' /* OPEN_ACCESS: open access\nOWN_MATERIAL: own material\nCOPYRIGHTED_MATERIAL: copyrighted material\nOTHER: other */,
    "copyright_classification_reason" VARCHAR(10000) NOT NULL,
    "item_type" VARCHAR(255) NOT NULL DEFAULT 'unknown' /* PRESENTATION: presentation\nREADER: reader\nBOOK: book\nARTICLE: article\nREPORT: report\nASSIGNMENT: assignment\nTHESIS: thesis\nMANUAL: manual\nUNKNOWN: unknown */,
    "item_type_classification_reason" VARCHAR(10000) NOT NULL,
    "pdf_name" VARCHAR(2048) NOT NULL,
    "publisher_name" VARCHAR(2048) NOT NULL,
    "copyright_holder" VARCHAR(2048) NOT NULL,
    "item_title" VARCHAR(2048) NOT NULL,
    "pdf_page_count" INT NOT NULL,
    "remarks" VARCHAR(10000) NOT NULL,
    "author_names" JSON NOT NULL,
    "doi" JSON NOT NULL,
    "isbn" JSON NOT NULL,
    "source_url" JSON NOT NULL,
    "license" JSON NOT NULL,
    "topic" JSON NOT NULL,
    "used_material_id" INT NOT NULL
);
CREATE TABLE IF NOT EXISTS "organization_data" (
    "created_at" TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    "modified_at" TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    "id" INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
    "hierarchy_level" INT NOT NULL,
    "name" VARCHAR(2048) NOT NULL,
    "abbreviation" VARCHAR(255) NOT NULL,
    "full_abbreviation" VARCHAR(2048) NOT NULL UNIQUE,
    "parent_organization_id" INT REFERENCES "organization_data" ("id") ON DELETE CASCADE,
    CONSTRAINT "uid_organizatio_name_82fcd1" UNIQUE ("name", "abbreviation")
);
CREATE INDEX "idx_organizatio_name_ad14b8" ON "organization_data" ("name");
CREATE INDEX "idx_organizatio_abbrevi_60e58f" ON "organization_data" ("abbreviation");
CREATE INDEX "idx_organizatio_full_ab_d01449" ON "organization_data" ("full_abbreviation");
CREATE TABLE IF NOT EXISTS "faculty" (
    "created_at" TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    "modified_at" TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    "id" INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
    "hierarchy_level" INT NOT NULL,
    "name" VARCHAR(2048) NOT NULL,
    "abbreviation" VARCHAR(255) NOT NULL UNIQUE,
    "full_abbreviation" VARCHAR(2048) NOT NULL UNIQUE,
    "parent_organization_id" INT REFERENCES "organization_data" ("id") ON DELETE CASCADE
);
CREATE INDEX "idx_faculty_name_eef624" ON "faculty" ("name");
CREATE INDEX "idx_faculty_abbrevi_f9ae1a" ON "faculty" ("abbreviation");
CREATE INDEX "idx_faculty_full_ab_dcd2c6" ON "faculty" ("full_abbreviation");
CREATE TABLE IF NOT EXISTS "copyright_data" (
    "created_at" TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    "modified_at" TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    "material_id" INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
    "period" VARCHAR(255) NOT NULL /* 2020_1A: 2020-1A\n2020_1B: 2020-1B\n2020_2A: 2020-2A\n2020_2B: 2020-2B\n2020_3: 2020-3\n2020_SEM1: 2020-SEM1\n2020_SEM2: 2020-SEM2\n2020_JAAR: 2020-JAAR\n2021_1A: 2021-1A\n2021_1B: 2021-1B\n2021_2A: 2021-2A\n2021_2B: 2021-2B\n2021_3: 2021-3\n2021_SEM1: 2021-SEM1\n2021_SEM2: 2021-SEM2\n2021_JAAR: 2021-JAAR\n2022_1A: 2022-1A\n2022_1B: 2022-1B\n2022_2A: 2022-2A\n2022_2B: 2022-2B\n2022_3: 2022-3\n2022_SEM1: 2022-SEM1\n2022_SEM2: 2022-SEM2\n2022_JAAR: 2022-JAAR\n2023_1A: 2023-1A\n2023_1B: 2023-1B\n2023_2A: 2023-2A\n2023_2B: 2023-2B\n2023_3: 2023-3\n2023_SEM1: 2023-SEM1\n2023_SEM2: 2023-SEM2\n2023_JAAR: 2023-JAAR\n2024_1A: 2024-1A\n2024_1B: 2024-1B\n2024_2A: 2024-2A\n2024_2B: 2024-2B\n2024_3: 2024-3\n2024_SEM1: 2024-SEM1\n2024_SEM2: 2024-SEM2\n2024_JAAR: 2024-JAAR\n2025_1A: 2025-1A\n2025_1B: 2025-1B\n2025_2A: 2025-2A\n2025_2B: 2025-2B\n2025_3: 2025-3\n2025_SEM1: 2025-SEM1\n2025_SEM2: 2025-SEM2\n2025_JAAR: 2025-JAAR\n2026_1A: 2026-1A\n2026_1B: 2026-1B\n2026_2A: 2026-2A\n2026_2B: 2026-2B\n2026_3: 2026-3\n2026_SEM1: 2026-SEM1\n2026_SEM2: 2026-SEM2\n2026_JAAR: 2026-JAAR\n2027_1A: 2027-1A\n2027_1B: 2027-1B\n2027_2A: 2027-2A\n2027_2B: 2027-2B\n2027_3: 2027-3\n2027_SEM1: 2027-SEM1\n2027_SEM2: 2027-SEM2\n2027_JAAR: 2027-JAAR\n2028_1A: 2028-1A\n2028_1B: 2028-1B\n2028_2A: 2028-2A\n2028_2B: 2028-2B\n2028_3: 2028-3\n2028_SEM1: 2028-SEM1\n2028_SEM2: 2028-SEM2\n2028_JAAR: 2028-JAAR\n2029_1A: 2029-1A\n2029_1B: 2029-1B\n2029_2A: 2029-2A\n2029_2B: 2029-2B\n2029_3: 2029-3\n2029_SEM1: 2029-SEM1\n2029_SEM2: 2029-SEM2\n2029_JAAR: 2029-JAAR\n2030_1A: 2030-1A\n2030_1B: 2030-1B\n2030_2A: 2030-2A\n2030_2B: 2030-2B\n2030_3: 2030-3\n2030_SEM1: 2030-SEM1\n2030_SEM2: 2030-SEM2\n2030_JAAR: 2030-JAAR */,
    "department" VARCHAR(2048) NOT NULL,
    "course_code" VARCHAR(255) NOT NULL,
    "course_name" VARCHAR(2048) NOT NULL,
    "url" VARCHAR(255) UNIQUE,
    "filename" VARCHAR(2048),
    "title" VARCHAR(2048),
    "owner" VARCHAR(2048),
    "filetype" VARCHAR(255) NOT NULL DEFAULT 'unknown' /* PDF: pdf\nPPT: ppt\nDOC: doc\nXLSX: xlsx\nMP4: mp4\nJPG: jpg\nPNG: png\nUNKNOWN: unknown\nFILE: file */,
    "classification" VARCHAR(255) NOT NULL DEFAULT 'lange overname' /* OPEN_ACCESS: open access\nKORTE_OVERNAME: korte overname\nMIDDELLANGE_OVERNAME: middellange overname\nLANGE_OVERNAME: lange overname\nEIGEN_MATERIAAL_POWERPOINT: eigen materiaal - powerpoint\nEIGEN_MATERIAAL_TITELINDICATIE: eigen materiaal - titelindicatie\nEIGEN_MATERIAAL_OVERIG: eigen materiaal - overig\nEIGEN_MATERIAAL: eigen materiaal\nONBEKEND: onbekend\nNIET_GEANALYSEERD: niet geanalyseerd\nIN_ONDERZOEK: in onderzoek\nVERWIJDERVERZOEK_VERSTUURD: verwijderverzoek verstuurd\nLICENTIE_BESCHIKBAAR: licentie beschikbaar */,
    "ml_prediction" VARCHAR(255) /* OPEN_ACCESS: open access\nKORTE_OVERNAME: korte overname\nMIDDELLANGE_OVERNAME: middellange overname\nLANGE_OVERNAME: lange overname\nEIGEN_MATERIAAL_POWERPOINT: eigen materiaal - powerpoint\nEIGEN_MATERIAAL_TITELINDICATIE: eigen materiaal - titelindicatie\nEIGEN_MATERIAAL_OVERIG: eigen materiaal - overig\nEIGEN_MATERIAAL: eigen materiaal\nONBEKEND: onbekend\nNIET_GEANALYSEERD: niet geanalyseerd\nIN_ONDERZOEK: in onderzoek\nVERWIJDERVERZOEK_VERSTUURD: verwijderverzoek verstuurd\nLICENTIE_BESCHIKBAAR: licentie beschikbaar */,
    "manual_classification" VARCHAR(2048),
    "manual_identifier" VARCHAR(2048),
    "scope" VARCHAR(255),
    "remarks" VARCHAR(10000),
    "auditor" VARCHAR(10000),
    "last_change" DATE,
    "status" VARCHAR(255) NOT NULL DEFAULT 'Published' /* PUBLISHED: Published\nUNPUBLISHED: Unpublished\nDELETED: Deleted */,
    "isbn" VARCHAR(255),
    "doi" VARCHAR(255),
    "in_collection" INT,
    "pagecount" INT NOT NULL,
    "wordcount" INT NOT NULL,
    "picturecount" INT NOT NULL,
    "author" VARCHAR(2048),
    "publisher" VARCHAR(2048),
    "reliability" INT NOT NULL,
    "pages_x_students" INT NOT NULL,
    "count_students_registered" INT NOT NULL,
    "retrieved_from_copyright_on" TIMESTAMP,
    "workflow_status" VARCHAR(255) NOT NULL DEFAULT 'ToDo' /* ToDo: ToDo\nDone: Done\nInProgress: InProgress */,
    "possible_fine" REAL,
    "infringement" VARCHAR(255) NOT NULL DEFAULT 'undetermined' /* YES: yes\nNO: no\nMAYBE: maybe\nUNDETERMINED: undetermined */,
    "faculty_id" VARCHAR(255) NOT NULL REFERENCES "faculty" ("abbreviation") ON DELETE CASCADE,
    "llm_classification_id" INT UNIQUE REFERENCES "llm_classification_data" ("id") ON DELETE CASCADE
);
CREATE INDEX "idx_copyright_d_departm_ad6f07" ON "copyright_data" ("department");
CREATE INDEX "idx_copyright_d_course__9d22fc" ON "copyright_data" ("course_code");
CREATE INDEX "idx_copyright_d_course__68968d" ON "copyright_data" ("course_name");
CREATE INDEX "idx_copyright_d_filenam_b3f28f" ON "copyright_data" ("filename");
CREATE INDEX "idx_copyright_d_ml_pred_382935" ON "copyright_data" ("ml_prediction");
CREATE INDEX "idx_copyright_d_manual__bb2fac" ON "copyright_data" ("manual_classification");
CREATE INDEX "idx_copyright_d_status_e5d6db" ON "copyright_data" ("status");
CREATE INDEX "idx_copyright_d_retriev_09f3f1" ON "copyright_data" ("retrieved_from_copyright_on");
CREATE INDEX "idx_copyright_d_workflo_4d93a6" ON "copyright_data" ("workflow_status");
CREATE TABLE IF NOT EXISTS "course_data" (
    "created_at" TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    "modified_at" TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    "cursuscode" INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
    "internal_id" INT NOT NULL UNIQUE,
    "year" INT NOT NULL,
    "name" VARCHAR(2048) NOT NULL,
    "short_name" VARCHAR(255),
    "ec" INT,
    "programme" VARCHAR(2048),
    "notes" VARCHAR(10000),
    "category" VARCHAR(2048),
    "faculty_id" VARCHAR(255) REFERENCES "faculty" ("abbreviation") ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS "person_data" (
    "created_at" TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    "modified_at" TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    "id" INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
    "input_name" VARCHAR(2048) NOT NULL UNIQUE,
    "main_name" VARCHAR(2048),
    "match_confidence" REAL,
    "first_name" VARCHAR(2048),
    "email" VARCHAR(2048),
    "people_page_url" VARCHAR(2048),
    "faculty_id" VARCHAR(255) REFERENCES "faculty" ("abbreviation") ON DELETE CASCADE
);
CREATE INDEX "idx_person_data_input_n_d53faf" ON "person_data" ("input_name");
CREATE TABLE IF NOT EXISTS "course_employee" (
    "created_at" TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    "modified_at" TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    "id" INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
    "role" VARCHAR(2048),
    "course_id" INT NOT NULL REFERENCES "course_data" ("cursuscode") ON DELETE CASCADE,
    "person_id" INT NOT NULL REFERENCES "person_data" ("id") ON DELETE CASCADE
);
CREATE TABLE IF NOT EXISTS "programme_data" (
    "id" INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL,
    "created_at" TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    "modified_at" TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    "cluster" VARCHAR(2048),
    "name" VARCHAR(2048) NOT NULL,
    "abbreviation" VARCHAR(255) NOT NULL,
    "programme_type" VARCHAR(255),
    "faculty_id" VARCHAR(255) REFERENCES "faculty" ("abbreviation") ON DELETE CASCADE,
    CONSTRAINT "uid_programme_d_name_60efd3" UNIQUE ("name", "abbreviation")
);
CREATE INDEX "idx_programme_d_name_a8877d" ON "programme_data" ("name");
CREATE INDEX "idx_programme_d_abbrevi_5cfb72" ON "programme_data" ("abbreviation");
CREATE TABLE IF NOT EXISTS "copyright_data_course_data" (
    "copyright_data_id" INT NOT NULL REFERENCES "copyright_data" ("material_id") ON DELETE CASCADE,
    "course_id" INT NOT NULL REFERENCES "course_data" ("cursuscode") ON DELETE CASCADE
);
CREATE UNIQUE INDEX "uidx_copyright_d_copyrig_c930bd" ON "copyright_data_course_data" ("copyright_data_id", "course_id");
CREATE TABLE IF NOT EXISTS "copyright_data_item_updates" (
    "copyright_data_id" INT NOT NULL REFERENCES "copyright_data" ("material_id") ON DELETE CASCADE,
    "itemupdate_id" INT NOT NULL REFERENCES "item_updates" ("id") ON DELETE CASCADE
);
CREATE UNIQUE INDEX "uidx_copyright_d_copyrig_42b84a" ON "copyright_data_item_updates" ("copyright_data_id", "itemupdate_id");
CREATE TABLE IF NOT EXISTS "person_data_organization_data" (
    "person_data_id" INT NOT NULL REFERENCES "person_data" ("id") ON DELETE CASCADE,
    "organization_id" INT NOT NULL REFERENCES "organization_data" ("id") ON DELETE CASCADE
);
CREATE UNIQUE INDEX "uidx_person_data_person__271921" ON "person_data_organization_data" ("person_data_id", "organization_id");
