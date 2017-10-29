CREATE TABLE [dbo].[BaseDatos] (
	[num_basedatos] [int] IDENTITY (1, 1) NOT NULL ,
	[nom_basedatos] [nvarchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[gls_coneccion] [nvarchar] (500) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[gls_formato_fecha] [nvarchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Cons_Perfil] (
	[num_perfil] [int] NOT NULL ,
	[num_consulta] [int] NOT NULL ,
	[fec_creacion] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Cons_Usuario] (
	[nom_usuario] [nvarchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[num_consulta] [int] NOT NULL ,
	[nom_creador] [nvarchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[fec_creacion] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Consultas] (
	[num_consulta] [int] IDENTITY (1, 1) NOT NULL ,
	[nom_consulta] [nvarchar] (132) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[num_basedatos] [int] NULL ,
	[gls_query] [text] NULL ,
	[nom_dueno] [nvarchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[nom_creador] [nvarchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[fec_creacion] [datetime] NULL ,
	[fec_ult_actualizacion] [datetime] NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Formatos] (
	[num_consulta] [int] NOT NULL ,
	[nom_columna] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[cod_tipo_dato_salida] [nvarchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ind_separador_miles] [nvarchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[num_decimales] [int] NULL ,
	[gls_formato_entrada] [nvarchar] (132) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[gls_formato_salida] [nvarchar] (132) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Log_Consultas] (
	[num_consulta] [int] NULL ,
	[nom_usuario] [nvarchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[fec_ejecucion] [datetime] NULL ,
	[hor_inicio] [nvarchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[gls_tiempo_utilizado] [nvarchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[num_registros] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Parametros] (
	[num_consulta] [int] NOT NULL ,
	[nom_parametro] [nvarchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[cod_tipo_dato] [nvarchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[cod_tipo_ayuda] [nvarchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[gls_ayuda_valores] [ntext] NULL ,
	[ind_opcional] [nvarchar] (1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[Perf_Usuario] (
	[nom_usuario] [nvarchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[num_perfil] [int] NOT NULL ,
	[fec_creacion] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Perfiles] (
	[num_perfil] [int] IDENTITY (1, 1) NOT NULL ,
	[nom_perfil] [nvarchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[fec_creacion] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Tipo_Datos] (
	[gls_proveedor] [nvarchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[num_tipo_columna_in] [int] NOT NULL ,
	[num_tipo_columna_out] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Tipo_Usuario] (
	[cod_tipo_usuario] [nvarchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ind_administrador] [char] (1) NULL ,
	[ind_crear_consultas] [char] (1) NULL ,
	[ind_autoasignar_consultas] [char] (1) NULL ,
	[ind_modificar_consultas] [char] (1) NULL ,
	[ind_eliminar_consultas] [char] (1) NULL ,
	[ind_ejecutar_consultas] [char] (1) NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Usuarios] (
	[nom_usuario] [nvarchar] (32) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[cod_tipo_usuario] [nvarchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[sysconfig] (
	[id_1] [nvarchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[id_2] [nvarchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[id_3] [nvarchar] (80) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[path_exe] [nvarchar] (132) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[path_hlp] [nvarchar] (132) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[BaseDatos] WITH NOCHECK ADD 
	CONSTRAINT [PK__BaseDatos__7E6CC920] PRIMARY KEY  CLUSTERED 
	(
		[num_basedatos]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Cons_Perfil] WITH NOCHECK ADD 
	CONSTRAINT [PK__Cons_Perfil__08EA5793] PRIMARY KEY  CLUSTERED 
	(
		[num_perfil],
		[num_consulta]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Cons_Usuario] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[nom_usuario],
		[num_consulta]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Consultas] WITH NOCHECK ADD 
	CONSTRAINT [PK__Consultas__03317E3D] PRIMARY KEY  CLUSTERED 
	(
		[num_consulta]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Formatos] WITH NOCHECK ADD 
	CONSTRAINT [PK_Formatos] PRIMARY KEY  CLUSTERED 
	(
		[num_consulta],
		[nom_columna]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Parametros] WITH NOCHECK ADD 
	CONSTRAINT [PK__Parametros__060DEAE8] PRIMARY KEY  CLUSTERED 
	(
		[num_consulta],
		[nom_parametro]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Perf_Usuario] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[nom_usuario],
		[num_perfil]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Perfiles] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[num_perfil]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Tipo_Datos] WITH NOCHECK ADD 
	CONSTRAINT [PK_Tipo_Datos] PRIMARY KEY  CLUSTERED 
	(
		[gls_proveedor],
		[num_tipo_columna_in]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Tipo_Usuario] WITH NOCHECK ADD 
	CONSTRAINT [PK__Tipo_Usuario__7C8480AE] PRIMARY KEY  CLUSTERED 
	(
		[cod_tipo_usuario]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Usuarios] WITH NOCHECK ADD 
	 PRIMARY KEY  CLUSTERED 
	(
		[nom_usuario]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[BaseDatos] ADD 
	CONSTRAINT [IX_BaseDatos] UNIQUE  NONCLUSTERED 
	(
		[nom_basedatos]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Consultas] ADD 
	CONSTRAINT [IX_Consultas] UNIQUE  NONCLUSTERED 
	(
		[nom_consulta]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Perfiles] ADD 
	CONSTRAINT [IX_Perfiles] UNIQUE  NONCLUSTERED 
	(
		[nom_perfil]
	)  ON [PRIMARY] 
GO

 CREATE  INDEX [Cons_Perfil_FKIndex1] ON [dbo].[Cons_Perfil]([num_perfil]) ON [PRIMARY]
GO

 CREATE  INDEX [Cons_Perfil_FKIndex2] ON [dbo].[Cons_Perfil]([num_consulta]) ON [PRIMARY]
GO

 CREATE  INDEX [Cons_Usuario_FKIndex1] ON [dbo].[Cons_Usuario]([nom_usuario]) ON [PRIMARY]
GO

 CREATE  INDEX [Cons_Usuario_FKIndex2] ON [dbo].[Cons_Usuario]([num_consulta]) ON [PRIMARY]
GO

 CREATE  INDEX [Consultas_FKIndex1] ON [dbo].[Consultas]([num_basedatos]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Log_Consultas] ON [dbo].[Log_Consultas]([num_consulta]) WITH  FILLFACTOR = 80 ON [PRIMARY]
GO

 CREATE  INDEX [Parametros_FKIndex1] ON [dbo].[Parametros]([num_consulta]) ON [PRIMARY]
GO

 CREATE  INDEX [Perf_Usuario_FKIndex1] ON [dbo].[Perf_Usuario]([nom_usuario]) ON [PRIMARY]
GO

 CREATE  INDEX [Perf_Usuario_FKIndex2] ON [dbo].[Perf_Usuario]([num_perfil]) ON [PRIMARY]
GO

 CREATE  INDEX [Usuarios_FKIndex1] ON [dbo].[Usuarios]([cod_tipo_usuario]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Cons_Perfil] ADD 
	CONSTRAINT [FK__Cons_Perf__num_c__0AD2A005] FOREIGN KEY 
	(
		[num_consulta]
	) REFERENCES [dbo].[Consultas] (
		[num_consulta]
	),
	CONSTRAINT [FK__Cons_Perf__num_p__09DE7BCC] FOREIGN KEY 
	(
		[num_perfil]
	) REFERENCES [dbo].[Perfiles] (
		[num_perfil]
	)
GO

ALTER TABLE [dbo].[Cons_Usuario] ADD 
	 FOREIGN KEY 
	(
		[nom_usuario]
	) REFERENCES [dbo].[Usuarios] (
		[nom_usuario]
	),
	CONSTRAINT [FK__Cons_Usua__num_c__1367E606] FOREIGN KEY 
	(
		[num_consulta]
	) REFERENCES [dbo].[Consultas] (
		[num_consulta]
	)
GO

ALTER TABLE [dbo].[Consultas] ADD 
	CONSTRAINT [FK__Consultas__num_b__0425A276] FOREIGN KEY 
	(
		[num_basedatos]
	) REFERENCES [dbo].[BaseDatos] (
		[num_basedatos]
	)
GO

ALTER TABLE [dbo].[Parametros] ADD 
	CONSTRAINT [FK__Parametro__num_c__07020F21] FOREIGN KEY 
	(
		[num_consulta]
	) REFERENCES [dbo].[Consultas] (
		[num_consulta]
	)
GO

ALTER TABLE [dbo].[Perf_Usuario] ADD 
	 FOREIGN KEY 
	(
		[nom_usuario]
	) REFERENCES [dbo].[Usuarios] (
		[nom_usuario]
	),
	 FOREIGN KEY 
	(
		[num_perfil]
	) REFERENCES [dbo].[Perfiles] (
		[num_perfil]
	)
GO

ALTER TABLE [dbo].[Usuarios] ADD 
	CONSTRAINT [FK__Usuarios__cod_ti__014935CB] FOREIGN KEY 
	(
		[cod_tipo_usuario]
	) REFERENCES [dbo].[Tipo_Usuario] (
		[cod_tipo_usuario]
	)
GO

