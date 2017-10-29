CREATE TABLE SAT_TIPO_USUA
(
    TUSU_ID         NUMBER(2,0)
   ,TUSU_NOMB       VARCHAR2(20)
   ,TUSU_ADMN       NUMBER(1,0)
   ,TUSU_CONS_CREA  NUMBER(1,0)
   ,TUSU_CONS_AASG  NUMBER(1,0)
   ,TUSU_CONS_MDIF  NUMBER(1,0)
   ,TUSU_CONS_ELIM  NUMBER(1,0)
   ,TUSU_CONS_EJEC  NUMBER(1,0)
)
    NOLOGGING
    TABLESPACE SATE_DATA_DAT
;

ALTER TABLE SAT_TIPO_USUA
    ADD CONSTRAINT TUSU_PK PRIMARY KEY(TUSU_ID)
    USING INDEX TABLESPACE SATE_DATA_IDX
;

ALTER TABLE SAT_TIPO_USUA
    ADD CONSTRAINT TUSU_UK UNIQUE (TUSU_NOMB)
    USING INDEX TABLESPACE SATE_DATA_IDX
;