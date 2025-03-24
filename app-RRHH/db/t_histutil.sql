CREATE TABLE "public"."utilidades_printhist" (
  "id" int4 NOT NULL DEFAULT nextval('utilidades_printhist_id_seq'::regclass),
  "codigo_empleado" varchar COLLATE "pg_catalog"."default",
  "name_utilidad" text COLLATE "pg_catalog"."default",
  "ci" varchar(11) COLLATE "pg_catalog"."default",
  "nombap" varchar(60) COLLATE "pg_catalog"."default",
  "devengado_util" numeric,
  "aporte_ss" numeric,
  "imp_ingp" numeric,
  "descuento_rm" numeric,
  "neto_cobrar" numeric,
  "area" varchar(100) COLLATE "pg_catalog"."default"
)
;

ALTER TABLE "public"."utilidades_printhist" 
  OWNER TO "postgres";