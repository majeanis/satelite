/* Crea Tipo de Usuario Administración */
insert into tipo_usuario
(cod_tipo_usuario
,ind_administrador 
,ind_crear_consultas 
,ind_autoasignar_consultas 
,ind_modificar_consultas 
,ind_eliminar_consultas 
,ind_ejecutar_consultas
)
values 
('ADMIN','S','S','S','S','S','S')

/* Crea Usuario */
insert into usuarios
(nom_usuario
,cod_tipo_usuario
)
values
('alabrin'
,'ADMIN'
)

/* Crea Tipo de Usuario de Consulta */

insert into tipo_usuario
(cod_tipo_usuario
,ind_administrador 
,ind_crear_consultas 
,ind_autoasignar_consultas 
,ind_modificar_consultas 
,ind_eliminar_consultas 
,ind_ejecutar_consultas
)
values 
('CONSULTA','N','N','N','N','N','S')

