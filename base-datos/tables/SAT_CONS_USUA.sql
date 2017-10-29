CREATE TABLE SAT_CONS_USUA
(
    CONS_ID         NUMBER(8,0)
   ,USUA_ID         NUMBER(8,0)
   ,AUDI_USUA_ID    NUMBER(6,0)
   ,AUDI_CREA       TIMESTAMP
)
    NOLOGGING
    TABLESPACE SATE_DATA_DAT
;

ALTER TABLE SAT_CONS_USUA
    ADD CONSTRAINT COUS_PK PRIMARY KEY (CONS_ID, USUA_ID)
    USING INDEX TABLESPACE SATE_DATA_IDX
;

CREATE INDEX COUS_CONS_FK ON SAT_CONS_USUA
(
    CONS_ID
)
    NOLOGGING
    TABLESPACE SATE_DATA_IDX
;

CREATE INDEX COUS_USUA_FK ON SAT_CONS_USUA
(
    USUA_ID
)
    NOLOGGING
    TABLESPACE SATE_DATA_IDX
;
