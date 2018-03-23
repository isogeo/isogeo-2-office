DROP TABLE IF EXISTS "metadata";
DROP TABLE IF EXISTS "md_types";
CREATE TABLE "md_types" ("id" integer NOT NULL PRIMARY KEY AUTOINCREMENT, "comment" text NOT NULL, "label" varchar(50) NOT NULL UNIQUE);
INSERT INTO "md_types" VALUES(1,'','VECTOR');
INSERT INTO "md_types" VALUES(2,'','RASTER');
INSERT INTO "md_types" VALUES(3,'','SERVICE');
INSERT INTO "md_types" VALUES(4,'','RESOURCE');
INSERT INTO "md_types" VALUES(5,'','SERIES');
CREATE TABLE "metadata" ("id" text NOT NULL PRIMARY KEY AUTOINCREMENT, "title" varchar(300) NOT NULL, "abstract" text NOT NULL, "md_dt_crea" datetime NOT NULL, "md_dt_shared_first" datetime NOT NULL, "md_dt_shared_last" datetime NOT NULL, "md_dt_update" datetime NOT NULL, "rs_dt_crea" datetime NOT NULL, "rs_dt_update" datetime NOT NULL, "md_type_id" integer NOT NULL REFERENCES "metadatatype" ("label"));
