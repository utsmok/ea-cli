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
    "allowed_usage" VARCHAR(255) NOT NULL DEFAULT 'undetermined',
    "allowed_usage_reasoning" VARCHAR(10000) NOT NULL,
    "copyright_status" VARCHAR(255) NOT NULL DEFAULT 'other',
    "copyright_classification_reason" VARCHAR(10000) NOT NULL,
    "item_type" VARCHAR(255) NOT NULL DEFAULT 'unknown',
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
    "period" VARCHAR(255) NOT NULL,
    "department" VARCHAR(2048) NOT NULL,
    "course_code" VARCHAR(255) NOT NULL,
    "course_name" VARCHAR(2048) NOT NULL,
    "url" VARCHAR(255) UNIQUE,
    "filename" VARCHAR(2048),
    "title" VARCHAR(2048),
    "owner" VARCHAR(2048),
    "filetype" VARCHAR(255) NOT NULL DEFAULT 'unknown',
    "classification" VARCHAR(255) NOT NULL DEFAULT 'lange overname',
    "ml_prediction" VARCHAR(255),
    "manual_classification" VARCHAR(2048),
    "manual_identifier" VARCHAR(2048),
    "scope" VARCHAR(255),
    "remarks" VARCHAR(10000),
    "auditor" VARCHAR(10000),
    "last_change" DATE,
    "status" VARCHAR(255) NOT NULL DEFAULT 'Published',
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
    "workflow_status" VARCHAR(255) NOT NULL DEFAULT 'ToDo',
    "possible_fine" REAL,
    "infringement" VARCHAR(255) NOT NULL DEFAULT 'undetermined',
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
