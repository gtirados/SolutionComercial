select * from NUMERACION_DOCUMENTOS

alter table numeracion_documentos add serie char(4)
update NUMERACION_DOCUMENTOS set serie = 'T001'

update TABLAS set TAB_CONTABLE2 ='01' where TAB_CODCIA ='00' and TAB_TIPREG= 40 and TAB_NUMTAB = 1
update TABLAS set tab_nomcorto =TAB_CONTABLE2 where TAB_CODCIA ='00' and TAB_TIPREG= 40