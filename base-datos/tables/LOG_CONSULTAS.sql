CREATE TABLE LOG_CONSULTAS
(
    NUM_LOG                 NUMBER(20,0)
   ,NUM_CONSULTA            NUMBER(8,0)
   ,NOM_USUARIO             VARCHAR2(50)
   ,FEC_EJECUCION           TIMESTAMP
   ,HOR_INICIO              TIMESTAMP
   ,GLS_TIEMPO_UTILIZADO    NUMBER(8,0)
   ,NUM_REGISTROS           NUMBER(8,0)
)
    NOLOGGING
    TABLESPACE SATE_DATA_DAT
;

ALTER TABLE LOG_CONSULTAS
    ADD CONSTRAINT LOGS_PK PRIMARY KEY (NUM_LOG)
    USING INDEX TABLESPACE SATE_DATA_IDX
;

CREATE INDEX LOGS_USUA_FK ON LOG_CONSULTAS
(
    NOM_USUARIO
)
    NOLOGGING
    TABLESPACE SATE_DATA_DAT
;

CREATE INDEX LOGS_CONS_FK ON LOG_CONSULTAS
(
    NUM_CONSULTA
)
    NOLOGGING
    TABLESPACE SATE_DATA_DAT
;


CREATE SEQUENCE LOG_CONSULTAS_ID
    START WITH 1
    INCREMENT BY 1
    MINVALUE 1
    NOMAXVALUE
    NOCACHE
    ORDER
;
