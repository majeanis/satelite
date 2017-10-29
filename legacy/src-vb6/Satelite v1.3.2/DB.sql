SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_GrabaUsuariosPorLote]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[usp_GrabaUsuariosPorLote]
	(@nom_user	nvarchar(32)  
	,@gls_usuarios	ntext  
	) as  
  
DECLARE  
	 @wl_num_error	integer   
	,@wl_gls_error	nvarchar(132)  
	,@wl_num_filas	integer  
	,@idoc		integer  
  
BEGIN  
	--<V1.3.0>
	-- Graba los usuarios asignados a un lote

	EXEC sp_xml_preparedocument @idoc OUTPUT, @gls_usuarios  
  
	SET NOCOUNT ON  
	BEGIN TRANSACTION  

	-- Inserta registros   
	INSERT INTO Lote_Usuario  
	      (nom_usuario  
	      ,num_lote  
	      ,fec_creacion  
	      )  
	SELECT nom_usuario  
	      ,num_lote  
	      ,getdate()  
	FROM OPENXML (@idoc, ''/ROOT/LoteUsuario'', 1)  
	WITH (num_lote     integer  
	     ,nom_usuario  nvarchar(32)  
	     )  
	
	SET @wl_num_error = @@ERROR  
	IF @wl_num_error <> 0  
		BEGIN  
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al insertar usuarios por lote''  
		GOTO HandError  
		END  
	
	COMMIT TRANSACTION  
	SET NOCOUNT OFF  
	
	EXEC sp_xml_removedocument @idoc  
	
	--</V1.3.0>
	RETURN  
  
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
END  
' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_EliminaLotesPorUsuario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[usp_EliminaLotesPorUsuario]
	(@nom_usuario nvarchar(32)  
	,@gls_lotes   ntext  
	) AS
	
DECLARE  
	 @wl_num_error	integer   
	,@wl_gls_error	nvarchar(132)  
	,@wl_num_filas	integer  
	,@idoc		integer  
	
BEGIN  
	--<V1.3.0>
	-- Elimina los lotes asignados a un usuario

	EXEC sp_xml_preparedocument @idoc OUTPUT, @gls_lotes  
	
	SET NOCOUNT ON  
	BEGIN TRANSACTION  
	
	-- Elimina lotes por usuario  
	DELETE FROM lote_usuario  
	WHERE nom_usuario = @nom_usuario  
	AND   num_lote IN  
	     (SELECT num_lote  
	      FROM OPENXML (@idoc, ''/ROOT/LoteUsuario'',1)  
	      WITH (num_lote    integer  
	           ,nom_usuario nvarchar(32))  
	     )  
	
	SET @wl_num_error = @@ERROR  
	IF @wl_num_error <> 0  
		BEGIN  
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al eliminar lotes por usuario''
		GOTO HandError  
	END  
	
	COMMIT TRANSACTION  
	SET NOCOUNT OFF  
	
	EXEC sp_xml_removedocument @idoc  
	
	RETURN  
	
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
	--</V1.3.0>
END  
' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_LotesPorUsuario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[usp_LotesPorUsuario]
	(@nom_usuario nvarchar(32)
	) AS  
DECLARE  
	 @wl_num_error	integer   
	,@wl_gls_error	nvarchar(132)  
BEGIN
	--<V1.3.0>
	-- Selecciona todos los lotes asignados al usuarios 

	SELECT l.*
	FROM lote_usuario lu  
	    ,lotes        l
	WHERE lu.nom_usuario = @nom_usuario  
	AND   l.num_lote     = lu.num_lote  
	ORDER BY l.nom_lote  

	SET @wl_num_error = @@ERROR  
	IF @wl_num_error <> 0  
		BEGIN  
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al consultar Lotes por Usuarios''  
		GOTO HandError  
	END  

	RETURN  
	
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
	--</V1.3.0>
END
' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Lotes]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[Lotes](
	[num_lote] [int] IDENTITY(1,1) NOT NULL,
	[nom_lote] [nvarchar](132) NULL,
	[nom_creador] [nvarchar](32) NULL,
	[nom_solicitante] [nvarchar](32) NULL,
	[fec_creacion] [datetime] NULL,
 CONSTRAINT [PK_Lotes] PRIMARY KEY CLUSTERED 
(
	[num_lote] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_UsuariosPorLote]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[usp_UsuariosPorLote]
	(@num_lote integer
	) AS  
DECLARE  
	 @wl_num_error	integer   
	,@wl_gls_error	nvarchar(132)  
BEGIN
	--<V1.3.0>
	-- Selecciona todos los usuarios asignados al lote

	SELECT lu.*
	FROM lotes        l  
	    ,lote_usuario lu  
	WHERE l.num_lote   = @num_lote  
	AND   lu.num_lote  = l.num_lote  
	ORDER BY lu.nom_usuario  

	SET @wl_num_error = @@ERROR  
	IF @wl_num_error <> 0  
		BEGIN  
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al consultar Usuarios por Lote''  
		GOTO HandError  
	END  

	RETURN  
	
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
	--</V1.3.0>
END
' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[BaseDatos]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[BaseDatos](
	[num_basedatos] [int] IDENTITY(1,1) NOT NULL,
	[nom_basedatos] [varchar](80) NULL,
	[gls_coneccion] [varchar](500) NULL,
	[gls_formato_fecha] [nvarchar](32) NULL,
 CONSTRAINT [PK__BaseDatos__7E6CC920] PRIMARY KEY CLUSTERED 
(
	[num_basedatos] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY],
 CONSTRAINT [IX_BaseDatos] UNIQUE NONCLUSTERED 
(
	[nom_basedatos] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Formatos]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[Formatos](
	[num_consulta] [int] NOT NULL,
	[nom_columna] [nvarchar](50) NOT NULL,
	[cod_tipo_dato_salida] [nvarchar](12) NULL,
	[ind_separador_miles] [nvarchar](1) NULL,
	[num_decimales] [int] NULL,
	[gls_formato_entrada] [nvarchar](132) NULL,
	[gls_formato_salida] [nvarchar](132) NULL,
 CONSTRAINT [PK_Formatos] PRIMARY KEY CLUSTERED 
(
	[num_consulta] ASC,
	[nom_columna] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Log_Consultas]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[Log_Consultas](
	[num_consulta] [int] NULL,
	[nom_usuario] [nvarchar](32) NULL,
	[fec_ejecucion] [datetime] NULL,
	[hor_inicio] [nvarchar](12) NULL,
	[gls_tiempo_utilizado] [nvarchar](12) NULL,
	[num_registros] [int] NULL
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Perfiles]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[Perfiles](
	[num_perfil] [int] IDENTITY(1,1) NOT NULL,
	[nom_perfil] [varchar](32) NULL,
	[fec_creacion] [datetime] NULL,
PRIMARY KEY CLUSTERED 
(
	[num_perfil] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY],
 CONSTRAINT [IX_Perfiles] UNIQUE NONCLUSTERED 
(
	[nom_perfil] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Tipo_Datos]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[Tipo_Datos](
	[gls_proveedor] [nvarchar](80) NOT NULL,
	[num_tipo_columna_in] [int] NOT NULL,
	[num_tipo_columna_out] [int] NULL,
 CONSTRAINT [PK_Tipo_Datos] PRIMARY KEY CLUSTERED 
(
	[gls_proveedor] ASC,
	[num_tipo_columna_in] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Tipo_Usuario]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[Tipo_Usuario](
	[cod_tipo_usuario] [varchar](12) NOT NULL,
	[ind_administrador] [char](1) NULL,
	[ind_crear_consultas] [char](1) NULL,
	[ind_autoasignar_consultas] [char](1) NULL,
	[ind_modificar_consultas] [char](1) NULL,
	[ind_eliminar_consultas] [char](1) NULL,
	[ind_ejecutar_consultas] [char](1) NULL,
 CONSTRAINT [PK__Tipo_Usuario__7C8480AE] PRIMARY KEY CLUSTERED 
(
	[cod_tipo_usuario] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[sysconfig]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[sysconfig](
	[id_1] [nvarchar](80) NULL,
	[id_2] [nvarchar](80) NULL,
	[id_3] [nvarchar](80) NULL,
	[path_exe] [nvarchar](132) NULL,
	[path_hlp] [nvarchar](132) NULL
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_LeeProveedores]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[usp_LeeProveedores] as
select *
from   Proveedores






' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Cons_Lote]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[Cons_Lote](
	[num_lote] [int] NOT NULL,
	[num_consulta] [int] NOT NULL,
	[fec_creacion] [datetime] NULL,
 CONSTRAINT [PK_Cons_Grupo] PRIMARY KEY CLUSTERED 
(
	[num_lote] ASC,
	[num_consulta] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Lote_Usuario]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[Lote_Usuario](
	[nom_usuario] [nvarchar](32) NOT NULL,
	[num_lote] [int] NOT NULL,
	[fec_creacion] [datetime] NULL,
 CONSTRAINT [PK_Grup_Usuario] PRIMARY KEY CLUSTERED 
(
	[nom_usuario] ASC,
	[num_lote] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Consultas]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[Consultas](
	[num_consulta] [int] IDENTITY(1,1) NOT NULL,
	[nom_consulta] [varchar](132) NULL,
	[num_basedatos] [int] NULL,
	[gls_query] [text] NULL,
	[nom_dueno] [nvarchar](32) NULL,
	[nom_creador] [nvarchar](32) NULL,
	[fec_creacion] [datetime] NULL,
	[fec_ult_actualizacion] [datetime] NULL,
	[gls_horario_ejecucion] [nvarchar](40) NULL,
	[gls_archivo_salida] [nvarchar](500) NULL,
	[nom_hoja_salida] [nvarchar](40) NULL,
 CONSTRAINT [PK__Consultas__03317E3D] PRIMARY KEY CLUSTERED 
(
	[num_consulta] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY],
 CONSTRAINT [IX_Consultas] UNIQUE NONCLUSTERED 
(
	[nom_consulta] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Cons_Perfil]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[Cons_Perfil](
	[num_perfil] [int] NOT NULL,
	[num_consulta] [int] NOT NULL,
	[fec_creacion] [datetime] NULL,
 CONSTRAINT [PK__Cons_Perfil__08EA5793] PRIMARY KEY CLUSTERED 
(
	[num_perfil] ASC,
	[num_consulta] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Cons_Usuario]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[Cons_Usuario](
	[nom_usuario] [varchar](32) NOT NULL,
	[num_consulta] [int] NOT NULL,
	[nom_creador] [nvarchar](32) NULL,
	[fec_creacion] [datetime] NULL,
PRIMARY KEY CLUSTERED 
(
	[nom_usuario] ASC,
	[num_consulta] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Parametros]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[Parametros](
	[num_consulta] [int] NOT NULL,
	[nom_parametro] [nvarchar](80) NOT NULL,
	[cod_tipo_dato] [nvarchar](12) NULL,
	[cod_tipo_ayuda] [nvarchar](12) NULL,
	[gls_ayuda_valores] [ntext] NULL,
	[ind_opcional] [nvarchar](1) NULL,
 CONSTRAINT [PK__Parametros__060DEAE8] PRIMARY KEY CLUSTERED 
(
	[num_consulta] ASC,
	[nom_parametro] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Perf_Usuario]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[Perf_Usuario](
	[nom_usuario] [varchar](32) NOT NULL,
	[num_perfil] [int] NOT NULL,
	[fec_creacion] [datetime] NULL,
PRIMARY KEY CLUSTERED 
(
	[nom_usuario] ASC,
	[num_perfil] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Usuarios]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[Usuarios](
	[nom_usuario] [varchar](32) NOT NULL,
	[cod_tipo_usuario] [varchar](12) NULL,
PRIMARY KEY CLUSTERED 
(
	[nom_usuario] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_LeeLotes]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[usp_LeeLotes] AS
DECLARE  
	 @wl_num_error	integer   
	,@wl_gls_error	nvarchar(132)  
BEGIN
	--<V1.3.0>
	-- Selecciona todos los lotes

	SELECT *  
	FROM   lotes
	ORDER BY nom_lote

	SET @wl_num_error = @@ERROR  
	IF @wl_num_error <> 0  
		BEGIN  
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al consultar Lotes''
		GOTO HandError  
	END  

	RETURN  
	
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
	--</V1.3.0>
END
' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_EliminaLote]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[usp_EliminaLote]
	(@num_lote integer) as
DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
BEGIN
	--<V1.3.0>
	-- Elimina el lote y las consultas asociadas
	SET NOCOUNT ON
	BEGIN TRANSACTION

	-- Elimina las consultas asociadas al lote
	DELETE FROM cons_lote
	WHERE  num_lote = @num_lote

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al eliminar consultas del lote''
		GOTO HandError
		END

	-- Elimina lote
	DELETE FROM lotes
	WHERE  num_lote = @num_lote

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al eliminar lote''
		GOTO HandError
		END

	COMMIT TRANSACTION
	SET NOCOUNT OFF

	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
	--</V1.3.0>
END
' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_LotesSinUsuario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
------------------------------------------------------------------

CREATE PROCEDURE [dbo].[usp_LotesSinUsuario]
	(@nom_usuario nvarchar(32)
	) AS
DECLARE  
	 @wl_num_error	integer   
	,@wl_gls_error	nvarchar(132)  
BEGIN
	--<V1.3.0>
	-- Selecciona todos los lotes que NO estan asignados al usuario

	SELECT *  
	FROM   lotes
	WHERE  num_lote NOT IN
	      (SELECT DISTINCT num_lote  
	       FROM  lote_usuario  
	       WHERE nom_usuario = @nom_usuario  
	      )  
	ORDER BY nom_lote

	SET @wl_num_error = @@ERROR  
	IF @wl_num_error <> 0  
		BEGIN  
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al consultar Lotes no asignados al Usuario''  
		GOTO HandError  
	END  

	RETURN  
	
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
	--</V1.3.0>
END
' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_GrabaLote]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[usp_GrabaLote]
	(@num_lote		integer
	,@nom_lote		nvarchar(132)
	,@nom_creador		nvarchar(32)
	,@nom_solicitante	nvarchar(32)
	,@wl_ind_asignar_lote	nvarchar(1)
	,@gls_xml		ntext
	) AS

DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
	,@wl_num_filas			integer
	,@idoc				integer

BEGIN
	--<V1.3.0>
	-- Crea o Actualiza el lote. Dentro del xml vienen las consultas asignadas al lote
	-- Se crea la asignación del Usuario Solicitante en caso que se indique desde el programa Cliente

	EXEC sp_xml_preparedocument @idoc OUTPUT, @gls_xml

	SET NOCOUNT ON
	BEGIN TRANSACTION
	
	-- Valida que el nombre de la lote sea único (aun cuando lo hace el indice, se valida para mejorar el mensaje de error)
	IF @num_lote = 0
		BEGIN

		SELECT @wl_num_filas = COUNT(*)
		FROM   Lotes
		WHERE  nom_lote = @nom_lote
	
		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al validar nombre del lote''
			GOTO HandError
			END
	
		IF @wl_num_filas > 0
			BEGIN
			SELECT @wl_gls_error = ''Nombre del lote ya existe. Intente con otro nombre''
			GOTO HandError
			END

		END
	/* END IF */

	-- Una vez validado el nombre del lote, crea o actualiza lote
	IF @num_lote = 0
		BEGIN

		-- Crea lote
		INSERT INTO Lotes
		      (nom_lote
		      ,nom_creador
		      ,nom_solicitante
		      ,fec_creacion
		      )
		VALUES
		      (@nom_lote
		      ,@nom_creador
		      ,@nom_solicitante
		      ,getdate()
		      )

		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al insertar lote''
			GOTO HandError
			END

		-- Obtiene el numero de lote asignado
		SELECT @num_lote = num_lote
		FROM   Lotes
		WHERE  nom_lote = @nom_lote
	
		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al obtener número de lote''
			GOTO HandError
			END
		END
	ELSE
		BEGIN

		UPDATE Lotes
		SET  nom_lote        = @nom_lote
		    ,nom_solicitante = @nom_solicitante
		WHERE num_lote = @num_lote

		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al actualizar lote''
			GOTO HandError
			END

		END
	/* END IF */

	-- Graba consultas del lote (las elimina previamente en caso que se está actualizando)
	DELETE FROM Cons_Lote
	WHERE num_lote = @num_lote

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al eliminar consultas del lote''
		GOTO HandError
		END

	INSERT INTO Cons_Lote
	      (num_lote    
	      ,num_consulta 
	      ,fec_creacion
	      )
	SELECT @num_lote
	      ,num_consulta
	      ,getdate()
	FROM OPENXML (@idoc, ''/ROOT/ConsLote'',1)
	WITH (num_lote     integer
	     ,num_consulta integer)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al crear consultas del lote''
		GOTO HandError
		END

	-- Asigna al usuario solicitante al lote, en caso que se indique
	IF @wl_ind_asignar_lote = ''S''
		BEGIN

		INSERT INTO Lote_Usuario
		      (nom_usuario
		      ,num_lote
		      ,fec_creacion
		      )
		VALUES
		      (@nom_solicitante
		      ,@num_lote
		      ,getdate()
		      ) 

		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al asignar usuario solicitante al lote''
			GOTO HandError
			END

		END
	/* END IF */

	COMMIT TRANSACTION
	SET NOCOUNT OFF

	EXEC sp_xml_removedocument @idoc
	
	-- Devuelve la información del lote creada o actualizada
	SELECT *
	FROM   lotes
	WHERE  num_lote = @num_lote

	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
	--</V1.3.0>
END
' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_EliminaUsuariosPorLote]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[usp_EliminaUsuariosPorLote]
	(@num_lote	integer  
	,@gls_usuarios	ntext  
	) as  
	
DECLARE  
	 @wl_num_error	integer   
	,@wl_gls_error	nvarchar(132)  
	,@wl_num_filas	integer  
	,@idoc		integer  
	
BEGIN  
	--<V1.3.0>
	-- Elimina los usuarios asignados a un lote

	EXEC sp_xml_preparedocument @idoc OUTPUT, @gls_usuarios  
	
	SET NOCOUNT ON  
	BEGIN TRANSACTION  
	
	-- Elimina usuario por lote  
	DELETE FROM Lote_Usuario  
	WHERE num_lote = @num_lote
	AND   nom_usuario IN  
	     (SELECT nom_usuario  
	      FROM OPENXML (@idoc, ''/ROOT/LoteUsuario'',1)  
	      WITH (num_lote    integer  
	           ,nom_usuario nvarchar(32))  
	     )  
	
	SET @wl_num_error = @@ERROR  
	IF @wl_num_error <> 0  
		BEGIN  
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al eliminar usuarios por lote''  
		GOTO HandError  
	END  
	
	COMMIT TRANSACTION  
	SET NOCOUNT OFF  
	
	EXEC sp_xml_removedocument @idoc  
	
	RETURN  
	
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
	--</V1.3.0>
END  

' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_EliminaBaseDatos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[usp_EliminaBaseDatos]
	(@num_basedatos integer) as
declare
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
begin
	SET NOCOUNT ON
	BEGIN TRANSACTION

	-- Elimina base de datos
	DELETE FROM BaseDatos
	WHERE  num_basedatos = @num_basedatos

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al eliminar base de datos''
		GOTO HandError
		END

	COMMIT TRANSACTION
	SET NOCOUNT OFF

	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN

end


' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_GrabaBaseDatos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE  PROCEDURE [dbo].[usp_GrabaBaseDatos]
	(@num_basedatos		integer
	,@nom_basedatos		nvarchar(80)
	,@gls_coneccion		nvarchar(500)
	,@gls_formato_fecha	nvarchar(32)
	) as

DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
	,@wl_num_filas			integer

BEGIN
	SET NOCOUNT ON
	BEGIN TRANSACTION

	-- Valida que el nombre de la base de datos sea único
	IF @num_basedatos = 0
		BEGIN

		SELECT @wl_num_filas = COUNT(*)
		FROM   basedatos
		WHERE  nom_basedatos = @nom_basedatos
	
		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al validar nombre de la base de datos''
			GOTO HandError
			END
	
		IF @wl_num_filas > 0
			BEGIN
			SELECT @wl_gls_error = ''Nombre de la base de datos ya existe. Intente con otro nombre''
			GOTO HandError
			END

		END
	/* END IF */
	
	-- Crea o actualiza base de datos
	IF @num_basedatos = 0
		BEGIN

		-- Crea base de datos
		INSERT INTO basedatos
			(nom_basedatos
			,gls_coneccion
			,gls_formato_fecha
			)
		VALUES  (@nom_basedatos
			,@gls_coneccion
			,@gls_formato_fecha
			)

		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al insertar base de datos''
			GOTO HandError
			END

		-- Obtiene el numero de la base de datos asignado
		SELECT @num_basedatos = num_basedatos
		FROM   basedatos
		WHERE  nom_basedatos = @nom_basedatos
	
		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al obtener número de la base de datos''
			GOTO HandError
			END

		END
	ELSE
		BEGIN

		-- Valida que no exista otra base de datos con el mismo nombre
		SELECT @wl_num_filas = COUNT(*)
		FROM   basedatos
		WHERE  nom_basedatos  = @nom_basedatos
		AND    num_basedatos <> @num_basedatos
	
		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al validar nombre de la base de datos''
			GOTO HandError
			END
	
		IF @wl_num_filas > 0
			BEGIN
			SELECT @wl_gls_error = ''Nombre de la base de datos ya existe. Intente con otro nombre''
			GOTO HandError
			END

		-- Actualiza informacion de la base de datos
		UPDATE basedatos
		SET   nom_basedatos     = @nom_basedatos
		     ,gls_coneccion     = @gls_coneccion
		     ,gls_formato_fecha = @gls_formato_fecha
		WHERE num_basedatos = @num_basedatos

		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al actualizar base de datos''
			GOTO HandError
			END

		END
	/* END IF */

	COMMIT TRANSACTION
	SET NOCOUNT OFF
	
	-- Devuelve la información de la consulta creada o actualizada
	SELECT *
	FROM   basedatos
	WHERE  num_basedatos = @num_basedatos

	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
END









' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_LeeBaseDatos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[usp_LeeBaseDatos] 
	(@num_basedatos	integer) as
begin
	select *
	from   BaseDatos
	where  num_basedatos = @num_basedatos
end


' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_LeeBasesDeDatos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[usp_LeeBasesDeDatos] as
begin
	select *
	from   BaseDatos
	order by num_basedatos
end


' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_LeeConsultas]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE procedure [dbo].[usp_LeeConsultas] as
begin
	select c.num_consulta 
	      ,c.nom_consulta                                                                                                                         
	      ,c.num_basedatos 
	      ,b.nom_basedatos
	      ,c.gls_query                                                                                                                                                                                                                                            
	      ,c.nom_dueno
	      ,c.nom_creador                      
	      ,c.fec_creacion                                           
	      ,c.fec_ult_actualizacion                                  
	from   consultas c
	      ,basedatos b
	where c.num_basedatos = b.num_basedatos
	order by c.num_consulta
end




' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_PerfilesPorConsulta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[usp_PerfilesPorConsulta]
	(@num_consulta integer) as
begin
	select distinct p.num_perfil, p.nom_perfil, cp.fec_creacion
	from  cons_perfil cp
	     ,perfiles    p
	where cp.num_consulta = @num_consulta
	and   p.num_perfil    = cp.num_perfil
	order by p.nom_perfil
end



' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_PerfilesSinConsulta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[usp_PerfilesSinConsulta]
	(@num_consulta integer) as
begin

	select distinct *
	from perfiles
	where num_perfil not in (
	select distinct num_perfil
	from cons_perfil
	where num_consulta = @num_consulta
	)
	order by nom_perfil

end


' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_GrabaConsultasPorPerfiles]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[usp_GrabaConsultasPorPerfiles]
	(@nom_user		nvarchar(32)
	,@gls_cons_perfil	ntext
	) as

DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
	,@wl_num_filas			integer
	,@idoc				integer

BEGIN
	EXEC sp_xml_preparedocument @idoc OUTPUT, @gls_cons_perfil

	SET NOCOUNT ON
	BEGIN TRANSACTION
	
	-- Inserta registros
	INSERT INTO cons_perfil
		(num_perfil
		,num_consulta
		,fec_creacion
		)
	SELECT	 num_perfil
		,num_consulta
		,getdate()
	FROM	OPENXML (@idoc, ''/ROOT/ConsPerfil'',1)
	WITH	(num_perfil		integer
		,num_consulta		integer
		)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al insertar consultas por perfiles''
		GOTO HandError
		END


	COMMIT TRANSACTION
	SET NOCOUNT OFF
	
	EXEC sp_xml_removedocument @idoc
	
	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
END












' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_EliminaConsultasPorPerfil]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[usp_EliminaConsultasPorPerfil]
	(@num_perfil	integer
	,@gls_consultas	ntext
	) as

DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
	,@wl_num_filas			integer
	,@idoc				integer

BEGIN
	EXEC sp_xml_preparedocument @idoc OUTPUT, @gls_consultas

	SET NOCOUNT ON
	BEGIN TRANSACTION
	
	-- Elimina consultas por perfil
	DELETE FROM cons_perfil
	WHERE num_perfil = @num_perfil
	AND   num_consulta IN
		(SELECT num_consulta
		 FROM	OPENXML (@idoc, ''/ROOT/ConsPerfil'',1)
		 WITH	(num_perfil		integer
			,num_consulta		integer)
		)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al eliminar consultas por perfil''
		GOTO HandError
		END


	COMMIT TRANSACTION
	SET NOCOUNT OFF

	EXEC sp_xml_removedocument @idoc
	
	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
END









' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_EliminaPerfilesPorConsulta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[usp_EliminaPerfilesPorConsulta]
	(@num_consulta	integer
	,@gls_perfiles	ntext
	) as

DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
	,@wl_num_filas			integer
	,@idoc				integer

BEGIN
	EXEC sp_xml_preparedocument @idoc OUTPUT, @gls_perfiles

	SET NOCOUNT ON
	BEGIN TRANSACTION
	
	-- Elimina perfiles por consulta
	DELETE FROM cons_perfil
	WHERE num_consulta = @num_consulta
	AND   num_perfil IN
		(SELECT num_perfil
		 FROM	OPENXML (@idoc, ''/ROOT/ConsPerfil'',1)
		 WITH	(num_perfil		integer
			,num_consulta		integer)
		)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al eliminar perfiles por consulta''
		GOTO HandError
		END


	COMMIT TRANSACTION
	SET NOCOUNT OFF
	
	EXEC sp_xml_removedocument @idoc
	
	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
END






select *
from cons_perfil


' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_ConsultasSinPerfil]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[usp_ConsultasSinPerfil]
	(@num_perfil integer) as
begin
	select *
	from consultas
	where num_consulta not in (
	select distinct num_consulta
	from cons_perfil
	where num_perfil   = @num_perfil
	)
	order by nom_consulta

end


' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_ConsultasPorPerfil]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create PROCEDURE [dbo].[usp_ConsultasPorPerfil]
	(@num_perfil integer) as
begin
	select c.num_consulta, c.nom_consulta, cp.fec_creacion
	from perfiles    p
	    ,cons_perfil cp
	    ,consultas   c
	where p.num_perfil   = @num_perfil
	and   cp.num_perfil  = p.num_perfil
	and   c.num_consulta = cp.num_consulta
	order by c.nom_consulta
end


' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_ConsultasPerfilPorUsuario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[usp_ConsultasPerfilPorUsuario]
(@nom_usuario nvarchar(32) ) as

select perfiles.num_perfil, perfiles.nom_perfil, consultas.num_consulta, consultas.nom_consulta
from perf_usuario
    ,perfiles
    ,cons_perfil
    ,consultas
where perf_usuario.nom_usuario = @nom_usuario
and   perfiles.num_perfil      = perf_usuario.num_perfil
and   cons_perfil.num_perfil   = perf_usuario.num_perfil
and   consultas.num_consulta   = cons_perfil.num_consulta
order by perfiles.nom_perfil, consultas.nom_consulta


' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_ConsultasPorUsuario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[usp_ConsultasPorUsuario] 
	(@nom_usuario nvarchar(32)) as
begin
	select c.num_consulta, c.nom_consulta, c.nom_dueno, cu.nom_creador, cu.fec_creacion
	from cons_usuario cu
	    ,consultas    c
	where cu.nom_usuario = @nom_usuario
	and   c.num_consulta = cu.num_consulta
	order by c.nom_consulta
end




' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_ConsultasSinUsuario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[usp_ConsultasSinUsuario]
	(@nom_usuario nvarchar(32)) as
begin

	select *
	from consultas
	where num_consulta not in (
	select distinct num_consulta
	from cons_usuario
	where nom_usuario = @nom_usuario
	)
	order by nom_consulta

end





' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_EliminaConsultasPorUsuario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[usp_EliminaConsultasPorUsuario]
	(@nom_usuario	nvarchar(32)
	,@gls_consultas	ntext
	) as

DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
	,@wl_num_filas			integer
	,@idoc				integer

BEGIN
	EXEC sp_xml_preparedocument @idoc OUTPUT, @gls_consultas

	SET NOCOUNT ON
	BEGIN TRANSACTION
	
	-- Elimina consultas por usuario
	DELETE FROM cons_usuario
	WHERE nom_usuario = @nom_usuario
	AND   num_consulta IN
		(SELECT num_consulta
		 FROM	OPENXML (@idoc, ''/ROOT/ConsUsuario'',1)
		 WITH	(num_consulta		integer
			,nom_usuario		nvarchar(32))
		)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al eliminar consultas por usuario''
		GOTO HandError
		END


	COMMIT TRANSACTION
	SET NOCOUNT OFF

	EXEC sp_xml_removedocument @idoc
	
	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
END







' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_GrabaConsultasPorUsuarios]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[usp_GrabaConsultasPorUsuarios]
	(@nom_user		nvarchar(32)
	,@gls_usuarios		ntext
	) as

DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
	,@wl_num_filas			integer
	,@idoc				integer

BEGIN
	EXEC sp_xml_preparedocument @idoc OUTPUT, @gls_usuarios

	SET NOCOUNT ON
	BEGIN TRANSACTION
	
	-- Inserta registros 
	INSERT INTO cons_usuario
		(nom_usuario
		,num_consulta
		,nom_creador
		,fec_creacion
		)
	SELECT	 nom_usuario
		,num_consulta
		,@nom_user
		,getdate()
	FROM	OPENXML (@idoc, ''/ROOT/ConsUsuario'',1)
	WITH	(num_consulta		integer
		,nom_usuario		nvarchar(32)
		)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al insertar consultas por usuarios''
		GOTO HandError
		END


	COMMIT TRANSACTION
	SET NOCOUNT OFF
	
	EXEC sp_xml_removedocument @idoc
	
	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
END












' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_EliminaUsuariosPorConsulta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[usp_EliminaUsuariosPorConsulta]
	(@num_consulta	integer
	,@gls_usuarios	ntext
	) as

DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
	,@wl_num_filas			integer
	,@idoc				integer

BEGIN
	EXEC sp_xml_preparedocument @idoc OUTPUT, @gls_usuarios

	SET NOCOUNT ON
	BEGIN TRANSACTION
	
	-- Elimina usuarios por consulta
	DELETE FROM cons_usuario
	WHERE num_consulta = @num_consulta
	AND   nom_usuario IN
		(SELECT nom_usuario
		 FROM	OPENXML (@idoc, ''/ROOT/ConsUsuario'',1)
		 WITH	(num_consulta		integer
			,nom_usuario		nvarchar(32))
		)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al eliminar usuarios por consulta''
		GOTO HandError
		END


	COMMIT TRANSACTION
	SET NOCOUNT OFF
	
	EXEC sp_xml_removedocument @idoc
	
	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
END





' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_UsuariosPorConsulta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[usp_UsuariosPorConsulta]
(@num_consulta integer) as
begin
select distinct *
from cons_usuario
where num_consulta = @num_consulta
order by nom_usuario
end



' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_UsuariosSinConsulta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE procedure [dbo].[usp_UsuariosSinConsulta]
	(@num_consulta integer) as
begin

	select distinct *
	from usuarios
	where nom_usuario not in (
	select distinct nom_usuario
	from cons_usuario
	where num_consulta = @num_consulta
	)
	order by nom_usuario

end





' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_EliminaConsulta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[usp_EliminaConsulta] 
	(@num_consulta integer) as
declare
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
begin
	SET NOCOUNT ON
	BEGIN TRANSACTION

	-- Elimina parámetros de la consulta
	DELETE FROM parametros
	WHERE  num_consulta = @num_consulta

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al eliminar parámetros de la consulta''
		GOTO HandError
		END

	-- Elimina la consulta
	DELETE FROM consultas
	WHERE  num_consulta = @num_consulta

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al eliminar la consulta''
		GOTO HandError
		END

	COMMIT TRANSACTION
	SET NOCOUNT OFF

	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN

end


' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_DetalleConsulta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[usp_DetalleConsulta]
(@num_consulta integer) as
begin
	select *
	from   consultas   
	where  num_consulta = @num_consulta
end


' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_LeeTodasConsultasPorLote]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[usp_LeeTodasConsultasPorLote]
	(@num_lote	integer
	) AS
BEGIN
	--<V1.3.0>
	-- Selecciona todas las consultas y aquellas que ya estan asociadas al lote las marca con una N

	SELECT	C.*
	       ,CASE WHEN ISNULL(cl.num_lote,0)=0 THEN ''N'' ELSE ''S'' END as ind_lote
	FROM	consultas c,
		cons_lote cl
	WHERE	cl.num_consulta =* c.num_consulta 
	AND	cl.num_lote     = @num_lote

	--</V1.3.0>
END


' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_LeeConsultasPorLote]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[usp_LeeConsultasPorLote]
	(@num_lote	integer
	) AS
BEGIN
	--<V1.3.0>
	-- Selecciona todas las consultas del lote

	SELECT	C.*
	FROM	consultas c,
		cons_lote cl
	WHERE	cl.num_consulta = c.num_consulta 
	AND	cl.num_lote     = @num_lote

	--</V1.3.0>
END

' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_GrabaFormatosConsulta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[usp_GrabaFormatosConsulta]
	(@num_consulta	integer
	,@gls_formatos		ntext
	) as

DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
	,@wl_num_filas			integer
	,@idoc				integer

BEGIN
	EXEC sp_xml_preparedocument @idoc OUTPUT, @gls_formatos

	SET NOCOUNT ON
	BEGIN TRANSACTION

	-- Elimina informacion anterior
	DELETE FROM 	Formatos
	WHERE num_consulta = @num_consulta

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al eliminar formato de columnas''
		GOTO HandError
		END
	
	-- Inserta registros
	INSERT INTO Formatos
		(num_consulta
		,nom_columna
		,cod_tipo_dato_salida
		,ind_separador_miles
		,num_decimales
		,gls_formato_entrada
		,gls_formato_salida
		)
	SELECT @num_consulta
		,nom_columna
		,cod_tipo_dato_salida
		,ind_separador_miles
		,num_decimales
		,gls_formato_entrada
		,gls_formato_salida
	FROM	OPENXML (@idoc, ''/ROOT/Formatos'',1)
	WITH	(nom_columna		nvarchar(50)
		,cod_tipo_dato_salida	nvarchar(12)
		,ind_separador_miles	nvarchar(1)
		,num_decimales		integer
		,gls_formato_entrada	nvarchar(132)
		,gls_formato_salida	nvarchar(132)
		)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al insertar formato de columnas''
		GOTO HandError
		END


	COMMIT TRANSACTION
	SET NOCOUNT OFF
	
	EXEC sp_xml_removedocument @idoc
	
	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
END
' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_LeeFormatos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[usp_LeeFormatos]
	(@num_consulta integer) as
begin
	select *
	from  formatos
	where num_consulta = @num_consulta
	order by nom_columna
end

' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_GrabaLogEjecucion]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[usp_GrabaLogEjecucion]
	(@num_consulta		integer
	,@nom_usuario		nvarchar(32)
	,@fec_ejecucion		nvarchar(12)
	,@hor_inicio		nvarchar(12)
	,@gls_tiempo_utilizado	nvarchar(12)
	,@num_registros		integer
	) as

BEGIN
	INSERT INTO Log_Consultas
		(num_consulta 
		,nom_usuario
		,fec_ejecucion
		,hor_inicio
		,gls_tiempo_utilizado
		,num_registros
		)
	VALUES
		(@num_consulta
		,@nom_usuario
		,@fec_ejecucion
		,@hor_inicio
		,@gls_tiempo_utilizado
		,@num_registros
		)
END


' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_DetalleParametro]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[usp_DetalleParametro]
(@num_consulta	integer
,@nom_parametro nvarchar(32)) as
begin
	select *
	from   parametros
	where  num_consulta  = @num_consulta
	and    nom_parametro = @nom_parametro
end


' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_EliminaPerfilesPorUsuario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[usp_EliminaPerfilesPorUsuario]
	(@nom_usuario	nvarchar(32)
	,@gls_perfiles	ntext
	) as

DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
	,@wl_num_filas			integer
	,@idoc				integer

BEGIN
	EXEC sp_xml_preparedocument @idoc OUTPUT, @gls_perfiles

	SET NOCOUNT ON
	BEGIN TRANSACTION
	
	-- Elimina perfiles por usuario
	DELETE FROM perf_usuario
	WHERE nom_usuario = @nom_usuario
	AND   num_perfil IN
		(SELECT num_perfil
		 FROM	OPENXML (@idoc, ''/ROOT/PerfUsuario'',1)
		 WITH	(num_perfil		integer
			,nom_usuario		nvarchar(32))
		)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al eliminar perfiles por usuario''
		GOTO HandError
		END


	COMMIT TRANSACTION
	SET NOCOUNT OFF
	
	EXEC sp_xml_removedocument @idoc
	
	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
END







select * from perf_usuario


' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_EliminaUsuariosPorPerfil]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[usp_EliminaUsuariosPorPerfil]
	(@num_perfil	integer
	,@gls_usuarios	ntext
	) as

DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
	,@wl_num_filas			integer
	,@idoc				integer

BEGIN
	EXEC sp_xml_preparedocument @idoc OUTPUT, @gls_usuarios

	SET NOCOUNT ON
	BEGIN TRANSACTION
	
	-- Elimina usuarios por perfil
	DELETE FROM perf_usuario
	WHERE num_perfil = @num_perfil
	AND   nom_usuario IN
		(SELECT nom_usuario
		 FROM	OPENXML (@idoc, ''/ROOT/PerfUsuario'',1)
		 WITH	(num_perfil		integer
			,nom_usuario		nvarchar(32))
		)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al eliminar usuarios por perfil''
		GOTO HandError
		END


	COMMIT TRANSACTION
	SET NOCOUNT OFF
	
	EXEC sp_xml_removedocument @idoc
	
	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
END



' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_GrabaPerfilesPorUsuarios]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[usp_GrabaPerfilesPorUsuarios]
	(@nom_user		nvarchar(32)
	,@gls_usuarios		ntext
	) as

DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
	,@wl_num_filas			integer
	,@idoc				integer

BEGIN
	EXEC sp_xml_preparedocument @idoc OUTPUT, @gls_usuarios

	SET NOCOUNT ON
	BEGIN TRANSACTION
	
	-- Inserta registros 
	INSERT INTO perf_usuario
		(nom_usuario
		,num_perfil
		,fec_creacion
		)
	SELECT	 nom_usuario
		,num_perfil
		,getdate()
	FROM	OPENXML (@idoc, ''/ROOT/PerfUsuario'',1)
	WITH	(num_perfil		integer
		,nom_usuario		nvarchar(32)
		)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al insertar perfiles por usuarios''
		GOTO HandError
		END


	COMMIT TRANSACTION
	SET NOCOUNT OFF
	
	EXEC sp_xml_removedocument @idoc
	
	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
END
















' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_PerfilesPorUsuario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[usp_PerfilesPorUsuario] 
	(@nom_usuario nvarchar(32)) as
begin
	select p.num_perfil, p.nom_perfil, pu.fec_creacion
	from perf_usuario pu
	    ,perfiles     p
	where pu.nom_usuario = @nom_usuario
	and   p.num_perfil   = pu.num_perfil
	order by p.nom_perfil
end




' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_UsuariosSinPerfil]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[usp_UsuariosSinPerfil]
	(@num_perfil integer) as
begin
	select *
	from usuarios
	where nom_usuario not in (
	select distinct nom_usuario
	from perf_usuario
	where num_perfil   = @num_perfil
	)
	order by nom_usuario

end






' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_UsuariosPorPerfil]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create PROCEDURE [dbo].[usp_UsuariosPorPerfil]
	(@num_perfil integer) as
begin
	select pu.nom_usuario, pu.fec_creacion
	from perfiles     p
	    ,perf_usuario pu
	where p.num_perfil   = @num_perfil
	and   pu.num_perfil  = p.num_perfil
	order by pu.nom_usuario
end




' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_PerfilesSinUsuario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[usp_PerfilesSinUsuario]
	(@nom_usuario nvarchar(32)) as
begin

	select *
	from perfiles
	where num_perfil not in (
	select distinct num_perfil
	from perf_usuario
	where nom_usuario = @nom_usuario
	)
	order by nom_perfil

end







' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_LeePerfiles]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[usp_LeePerfiles] as
select *
from perfiles
order by num_perfil


' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_GrabaPerfil]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[usp_GrabaPerfil]
	(@num_perfil		integer
	,@nom_perfil		nvarchar(132)
	) as

DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
	,@wl_num_filas			integer

BEGIN
	SET NOCOUNT ON
	BEGIN TRANSACTION
	
	-- Valida que el nombre del perfil sea único
	IF @num_perfil = 0
		BEGIN

		SELECT @wl_num_filas = COUNT(*)
		FROM   perfiles
		WHERE  nom_perfil = @nom_perfil
	
		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al validar nombre del perfil''
			GOTO HandError
			END
	
		IF @wl_num_filas > 0
			BEGIN
			SELECT @wl_gls_error = ''Nombre del perfil ya existe. Intente con otro nombre''
			GOTO HandError
			END

		END
	/* END IF */

	-- Crea o actualiza perfil
	IF @num_perfil = 0
		BEGIN

		-- Crea perfil
		INSERT INTO perfiles
			(nom_perfil
			,fec_creacion
			)
		VALUES
			(@nom_perfil
			,getdate()
			)

		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al insertar perfil''
			GOTO HandError
			END

		-- Obtiene el numero del perfil asignado
		SELECT @num_perfil = num_perfil
		FROM   perfiles
		WHERE  nom_perfil = @nom_perfil
	
		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al obtener número del perfil''
			GOTO HandError
			END

		END
	ELSE
		BEGIN

		UPDATE perfiles
		SET   nom_perfil = @nom_perfil
		WHERE num_perfil = @num_perfil

		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al actualizar perfil''
			GOTO HandError
			END

		END
	/* END IF */

	COMMIT TRANSACTION
	SET NOCOUNT OFF
	
	-- Devuelve la información de la consulta creada o actualizada
	SELECT *
	FROM   perfiles
	WHERE  num_perfil = @num_perfil

	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
END






' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_EliminaPerfil]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[usp_EliminaPerfil]
	(@num_perfil integer) as
declare
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
begin
	SET NOCOUNT ON
	BEGIN TRANSACTION

	-- Elimina perfil
	DELETE FROM perfiles
	WHERE  num_perfil = @num_perfil

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al eliminar perfil''
		GOTO HandError
		END

	COMMIT TRANSACTION
	SET NOCOUNT OFF

	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN

end






' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_LeeTipoDatos]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[usp_LeeTipoDatos] AS
BEGIN
	select *
	from  tipo_datos
	order by gls_proveedor, num_tipo_columna_in
END



' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_LeeTipoUsuario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE procedure [dbo].[usp_LeeTipoUsuario]
	(@cod_tipo_usuario	nvarchar(12)) as
begin
	select *
	from   tipo_usuario
	where  cod_tipo_usuario = @cod_tipo_usuario
end


' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_LeeTiposUsuarios]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[usp_LeeTiposUsuarios] as
begin
	select *
	from   tipo_usuario
	order by cod_tipo_usuario

end

' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_LeeUsuarios]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE procedure [dbo].[usp_LeeUsuarios] as
	select usuarios.nom_usuario
	      ,usuarios.cod_tipo_usuario
	      ,tipo_usuario.ind_administrador
	      ,tipo_usuario.ind_crear_consultas
	      ,tipo_usuario.ind_modificar_consultas
	      ,tipo_usuario.ind_eliminar_consultas
	      ,tipo_usuario.ind_ejecutar_consultas
	from usuarios
	    ,tipo_usuario
	where usuarios.cod_tipo_usuario = tipo_usuario.cod_tipo_usuario
	order by usuarios.nom_usuario
' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_EliminaTipoUsuario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE procedure [dbo].[usp_EliminaTipoUsuario]
	(@cod_tipo_usuario nvarchar(12)) as
declare
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
begin
	SET NOCOUNT ON
	BEGIN TRANSACTION

	-- Elimina tipo usuario
	DELETE FROM tipo_usuario
	WHERE  cod_tipo_usuario = @cod_tipo_usuario

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al eliminar tipo usuario''
		GOTO HandError
		END

	COMMIT TRANSACTION
	SET NOCOUNT OFF

	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN

end


' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_DetalleUsuario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE procedure [dbo].[usp_DetalleUsuario]
(@nom_usuario nvarchar(32) ) as

	select tipo_usuario.cod_tipo_usuario 
	      ,tipo_usuario.ind_administrador
	      ,tipo_usuario.ind_crear_consultas
	      ,tipo_usuario.ind_modificar_consultas
	      ,tipo_usuario.ind_eliminar_consultas
	      ,tipo_usuario.ind_ejecutar_consultas
	from usuarios
	    ,tipo_usuario
	where usuarios.nom_usuario      = @nom_usuario
	and   usuarios.cod_tipo_usuario = tipo_usuario.cod_tipo_usuario
' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_GrabaTipoUsuario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[usp_GrabaTipoUsuario]
	(@cod_tipo_usuario	nvarchar(12)
	,@gls_campos		ntext
	,@cod_tipo_accion	nvarchar(08) 
	) as

DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
	,@wl_num_filas			integer
	,@idoc				integer

BEGIN
	EXEC sp_xml_preparedocument @idoc OUTPUT, @gls_campos

	SET NOCOUNT ON
	BEGIN TRANSACTION
	
	-- Valida que el nombre del tipo de usuario sea único
	IF @cod_tipo_accion = ''INS''
		BEGIN

		SELECT @wl_num_filas = COUNT(*)
		FROM   tipo_usuario
		WHERE  cod_tipo_usuario = @cod_tipo_usuario
	
		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al validar nombre del tipo de usuario''
			GOTO HandError
			END
	
		IF @wl_num_filas > 0
			BEGIN
			SELECT @wl_gls_error = ''Nombre del tipo de usuario ya existe. Intente con otro nombre''
			GOTO HandError
			END

		-- Crea usuario
		INSERT INTO tipo_usuario
			(cod_tipo_usuario 
			,ind_administrador 
			,ind_crear_consultas 
			,ind_autoasignar_consultas 
			,ind_modificar_consultas 
			,ind_eliminar_consultas 
			,ind_ejecutar_consultas 
			)
		SELECT	 @cod_tipo_usuario 
			,ind_administrador 
			,ind_crear_consultas 
			,ind_autoasignar_consultas 
			,ind_modificar_consultas 
			,ind_eliminar_consultas 
			,ind_ejecutar_consultas 
		FROM	OPENXML (@idoc, ''/ROOT/TipoUsuario'',1)
		WITH	(ind_administrador		nvarchar(1)
			,ind_crear_consultas 		nvarchar(1)
			,ind_autoasignar_consultas 	nvarchar(1)
			,ind_modificar_consultas 	nvarchar(1)
			,ind_eliminar_consultas 	nvarchar(1)
			,ind_ejecutar_consultas 	nvarchar(1)
			)

		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al insertar tipo de usuario''
			GOTO HandError
			END

		END
	ELSE
		BEGIN

		SELECT	 cod_tipo_usuario = convert(nvarchar(12), @cod_tipo_usuario )
			,ind_administrador 
			,ind_crear_consultas 
			,ind_autoasignar_consultas 
			,ind_modificar_consultas 
			,ind_eliminar_consultas 
			,ind_ejecutar_consultas 
		INTO    #tmp_tipo_usuario
		FROM	OPENXML (@idoc, ''/ROOT/TipoUsuario'',1)
		WITH	(ind_administrador		nvarchar(1)
			,ind_crear_consultas 		nvarchar(1)
			,ind_autoasignar_consultas 	nvarchar(1)
			,ind_modificar_consultas 	nvarchar(1)
			,ind_eliminar_consultas 	nvarchar(1)
			,ind_ejecutar_consultas 	nvarchar(1)
			)

		UPDATE tipo_usuario
		SET	 ind_administrador         = tmp.ind_administrador 
			,ind_crear_consultas       = tmp.ind_crear_consultas
			,ind_autoasignar_consultas = tmp.ind_autoasignar_consultas
			,ind_modificar_consultas   = tmp.ind_modificar_consultas
			,ind_eliminar_consultas    = tmp.ind_eliminar_consultas
			,ind_ejecutar_consultas    = tmp.ind_ejecutar_consultas 
		FROM   #tmp_tipo_usuario tmp
		WHERE  tipo_usuario.cod_tipo_usuario = @cod_tipo_usuario
		AND    tmp.cod_tipo_usuario          = @cod_tipo_usuario

		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al actualizar tipo de usuario''
			GOTO HandError
			END

		DROP TABLE #tmp_tipo_usuario

		END
	/* END IF */

	COMMIT TRANSACTION
	SET NOCOUNT OFF
	
	EXEC sp_xml_removedocument @idoc
	
	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
END
' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_GrabaUsuario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROCEDURE [dbo].[usp_GrabaUsuario]
	(@nom_usuario		nvarchar(32)
	,@cod_tipo_usuario	nvarchar(32)
	,@cod_tipo_accion	nvarchar(08) 
	) as

DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
	,@wl_num_filas			integer

BEGIN
	SET NOCOUNT ON
	BEGIN TRANSACTION
	
	-- Valida que el nombre de la consulta sea único (aun cuando lo hace el indice, se valida para mejorar el mensaje de error)
	IF @cod_tipo_accion = ''INS''
		BEGIN

		SELECT @wl_num_filas = COUNT(*)
		FROM   usuarios
		WHERE  nom_usuario = @nom_usuario
	
		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al validar nombre del usuario''
			GOTO HandError
			END
	
		IF @wl_num_filas > 0
			BEGIN
			SELECT @wl_gls_error = ''Nombre del usuario ya existe. Intente con otro nombre''
			GOTO HandError
			END

		-- Crea usuario
		INSERT INTO usuarios
			(nom_usuario
			,cod_tipo_usuario
			)
		VALUES
			(@nom_usuario
			,@cod_tipo_usuario
			)

		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al insertar usuario''
			GOTO HandError
			END

		END
	ELSE
		BEGIN

		UPDATE usuarios
		SET cod_tipo_usuario = @cod_tipo_usuario
		WHERE nom_usuario = @nom_usuario

		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al actualizar usuario''
			GOTO HandError
			END

		END
	/* END IF */

	COMMIT TRANSACTION
	SET NOCOUNT OFF
	
	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
END






' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_EliminaUsuario]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[usp_EliminaUsuario]
	(@nom_usuario nvarchar(32)) as
declare
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
begin
	SET NOCOUNT ON
	BEGIN TRANSACTION

	-- Elimina usuario
	DELETE FROM usuarios
	WHERE  nom_usuario = @nom_usuario

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al eliminar usuario''
		GOTO HandError
		END

	COMMIT TRANSACTION
	SET NOCOUNT OFF

	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN

end




' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_UsuariosSinLote]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[usp_UsuariosSinLote]
	(@num_lote integer
	) AS
DECLARE  
	 @wl_num_error	integer   
	,@wl_gls_error	nvarchar(132)  
BEGIN
	--<V1.3.0>
	-- Selecciona todos los usuarios que NO estan asignados al lote

	SELECT *  
	FROM  usuarios  
	WHERE nom_usuario NOT IN 
	     (SELECT DISTINCT nom_usuario  
	      FROM   lote_usuario  
	      WHERE  num_lote = @num_lote  
	     )
	ORDER BY nom_usuario  

	SET @wl_num_error = @@ERROR  
	IF @wl_num_error <> 0  
		BEGIN  
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al consultar Usuarios no asignados al Lote''  
		GOTO HandError  
	END  

	RETURN  
	
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
	--</V1.3.0>
END
' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_LeeSysConfig]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
create procedure [dbo].[usp_LeeSysConfig]
as
	select *
	from sysconfig

' 
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[usp_GrabaConsulta]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'CREATE PROCEDURE [dbo].[usp_GrabaConsulta]
	(@num_consulta			integer
	,@nom_consulta			nvarchar(132)
	,@num_basedatos			integer
	,@gls_query			ntext
	,@gls_parametros		ntext
	,@gls_formatos			ntext
	,@gls_horario_ejecucion		nvarchar(40)
	,@gls_archivo_salida		nvarchar(500)
	,@nom_hoja_salida		nvarchar(40)
	,@nom_user			nvarchar(32)
	,@nom_user_real			nvarchar(32)
	) as

DECLARE
	 @wl_ind_asignar_consulta	char(1)
	,@wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
	,@wl_num_filas			integer
	,@idoc				integer

BEGIN
	EXEC sp_xml_preparedocument @idoc OUTPUT, @gls_parametros

	SET NOCOUNT ON
	BEGIN TRANSACTION
	
	-- Valida que el nombre de la consulta sea único (aun cuando lo hace el indice, se valida para mejorar el mensaje de error)
	IF @num_consulta = 0
		BEGIN

		SELECT @wl_num_filas = COUNT(*)
		FROM   consultas
		WHERE  nom_consulta = @nom_consulta
	
		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al validar nombre de la consulta''
			GOTO HandError
			END
	
		IF @wl_num_filas > 0
			BEGIN
			SELECT @wl_gls_error = ''Nombre de la consulta ya existe. Intente con otro nombre''
			GOTO HandError
			END

		END
	/* END IF */

	-- Determina tipo de usuario, en nombre del usuario que crea la consulta
	SELECT @wl_ind_asignar_consulta = ind_autoasignar_consultas
	FROM  usuarios     u 
	     ,tipo_usuario tu
	WHERE u.nom_usuario       = @nom_user
	AND   tu.cod_tipo_usuario = u.cod_tipo_usuario

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al obtener detalle de usuario''
		GOTO HandError
		END
	
	-- Crea o actualiza consulta
	IF @num_consulta = 0
		BEGIN

		-- Crea consulta
		INSERT INTO consultas
			(nom_consulta
			,num_basedatos
			,gls_query
			,nom_dueno
			,nom_creador
			,fec_creacion
			,fec_ult_actualizacion
			,gls_horario_ejecucion
			--<V1.3.0>
			,gls_archivo_salida
			,nom_hoja_salida
			--</V1.3.0>
			)
		VALUES
			(@nom_consulta
			,@num_basedatos
			,@gls_query
			,@nom_user
			,@nom_user_real
			,getdate()
			,getdate()
			,@gls_horario_ejecucion
			--<V1.3.0>
			,@gls_archivo_salida
			,@nom_hoja_salida
			--</V1.3.0>
			)

		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al insertar consulta''
			GOTO HandError
			END

		-- Obtiene el numero de consulta asignado
		SELECT @num_consulta = num_consulta
		FROM   consultas
		WHERE  nom_consulta = @nom_consulta
	
		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al obtener número de consulta''
			GOTO HandError
			END

		-- Si el perfil, asigna automáticamente la cosulta al dueño, ingresa la consulta al usuario
		IF @wl_ind_asignar_consulta = ''S''
			BEGIN

			-- Asigna automáticamente la consulta al usuario
			INSERT INTO cons_usuario
				(nom_usuario
				,num_consulta
				,nom_creador
				,fec_creacion
				)
			VALUES
				(@nom_user
				,@num_consulta
				,@nom_user_real
				,getdate()
				)
			END

			SET @wl_num_error = @@ERROR
			IF @wl_num_error <> 0
				BEGIN
				SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al obtener detalle de usuario''
				GOTO HandError
				END
		/* end if */
		END
	ELSE
		BEGIN

		UPDATE consultas
		SET nom_consulta            = @nom_consulta
		   ,num_basedatos           = @num_basedatos
		   ,gls_query               = @gls_query
		   ,fec_ult_actualizacion   = getdate()
		   ,gls_horario_ejecucion   = @gls_horario_ejecucion
		   --<V1.3.0>
		   ,gls_archivo_salida      = @gls_archivo_salida
		   ,nom_hoja_salida         = @nom_hoja_salida
		   --</V1.3.0>
		WHERE num_consulta = @num_consulta

		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al actualizar consulta''
			GOTO HandError
			END

		END
	/* END IF */

	-- Graba parámetros de la consulta
	DELETE FROM parametros
	WHERE num_consulta = @num_consulta

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al eliminar parametros''
		GOTO HandError
		END

	INSERT INTO parametros
	SELECT	 @num_consulta
		,nom_parametro
		,cod_tipo_dato
		,cod_tipo_ayuda
		,gls_ayuda_valores
		,ind_opcional
	FROM	OPENXML (@idoc, ''/ROOT/Parametros'',1)
	WITH	(num_consulta		integer
		,nom_parametro		nvarchar(80)
		,cod_tipo_dato		nvarchar(12)
		,cod_tipo_ayuda		nvarchar(12)
		,gls_ayuda_valores	ntext
		,ind_opcional		nvarchar(1))

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al crear parámetros''
		GOTO HandError
		END

	-- Graba Formatos de la consulta
	EXEC usp_GrabaFormatosConsulta @num_consulta, @gls_formatos

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + '', error al grabar formatos''
		GOTO HandError
		END


	COMMIT TRANSACTION
	SET NOCOUNT OFF

	EXEC sp_xml_removedocument @idoc
	
	-- Devuelve la información de la consulta creada o actualizada
	SELECT *
	FROM   consultas
	WHERE  num_consulta = @num_consulta

	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
END


' 
END
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_Cons_Lote_Consultas]') AND parent_object_id = OBJECT_ID(N'[dbo].[Cons_Lote]'))
ALTER TABLE [dbo].[Cons_Lote]  WITH NOCHECK ADD  CONSTRAINT [FK_Cons_Lote_Consultas] FOREIGN KEY([num_consulta])
REFERENCES [dbo].[Consultas] ([num_consulta])
GO
ALTER TABLE [dbo].[Cons_Lote] CHECK CONSTRAINT [FK_Cons_Lote_Consultas]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_Cons_Lote_Lotes]') AND parent_object_id = OBJECT_ID(N'[dbo].[Cons_Lote]'))
ALTER TABLE [dbo].[Cons_Lote]  WITH NOCHECK ADD  CONSTRAINT [FK_Cons_Lote_Lotes] FOREIGN KEY([num_lote])
REFERENCES [dbo].[Lotes] ([num_lote])
GO
ALTER TABLE [dbo].[Cons_Lote] CHECK CONSTRAINT [FK_Cons_Lote_Lotes]
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK_Lote_Usuario_Lotes]') AND parent_object_id = OBJECT_ID(N'[dbo].[Lote_Usuario]'))
ALTER TABLE [dbo].[Lote_Usuario]  WITH CHECK ADD  CONSTRAINT [FK_Lote_Usuario_Lotes] FOREIGN KEY([num_lote])
REFERENCES [dbo].[Lotes] ([num_lote])
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK__Consultas__num_b__0425A276]') AND parent_object_id = OBJECT_ID(N'[dbo].[Consultas]'))
ALTER TABLE [dbo].[Consultas]  WITH CHECK ADD  CONSTRAINT [FK__Consultas__num_b__0425A276] FOREIGN KEY([num_basedatos])
REFERENCES [dbo].[BaseDatos] ([num_basedatos])
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK__Cons_Perf__num_c__0AD2A005]') AND parent_object_id = OBJECT_ID(N'[dbo].[Cons_Perfil]'))
ALTER TABLE [dbo].[Cons_Perfil]  WITH CHECK ADD  CONSTRAINT [FK__Cons_Perf__num_c__0AD2A005] FOREIGN KEY([num_consulta])
REFERENCES [dbo].[Consultas] ([num_consulta])
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK__Cons_Perf__num_p__09DE7BCC]') AND parent_object_id = OBJECT_ID(N'[dbo].[Cons_Perfil]'))
ALTER TABLE [dbo].[Cons_Perfil]  WITH CHECK ADD  CONSTRAINT [FK__Cons_Perf__num_p__09DE7BCC] FOREIGN KEY([num_perfil])
REFERENCES [dbo].[Perfiles] ([num_perfil])
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK__Cons_Usua__nom_u__489AC854]') AND parent_object_id = OBJECT_ID(N'[dbo].[Cons_Usuario]'))
ALTER TABLE [dbo].[Cons_Usuario]  WITH CHECK ADD FOREIGN KEY([nom_usuario])
REFERENCES [dbo].[Usuarios] ([nom_usuario])
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK__Cons_Usua__num_c__1367E606]') AND parent_object_id = OBJECT_ID(N'[dbo].[Cons_Usuario]'))
ALTER TABLE [dbo].[Cons_Usuario]  WITH CHECK ADD  CONSTRAINT [FK__Cons_Usua__num_c__1367E606] FOREIGN KEY([num_consulta])
REFERENCES [dbo].[Consultas] ([num_consulta])
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK__Parametro__num_c__07020F21]') AND parent_object_id = OBJECT_ID(N'[dbo].[Parametros]'))
ALTER TABLE [dbo].[Parametros]  WITH CHECK ADD  CONSTRAINT [FK__Parametro__num_c__07020F21] FOREIGN KEY([num_consulta])
REFERENCES [dbo].[Consultas] ([num_consulta])
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK__Perf_Usua__nom_u__4C6B5938]') AND parent_object_id = OBJECT_ID(N'[dbo].[Perf_Usuario]'))
ALTER TABLE [dbo].[Perf_Usuario]  WITH CHECK ADD FOREIGN KEY([nom_usuario])
REFERENCES [dbo].[Usuarios] ([nom_usuario])
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK__Perf_Usua__num_p__4D5F7D71]') AND parent_object_id = OBJECT_ID(N'[dbo].[Perf_Usuario]'))
ALTER TABLE [dbo].[Perf_Usuario]  WITH CHECK ADD FOREIGN KEY([num_perfil])
REFERENCES [dbo].[Perfiles] ([num_perfil])
GO
IF NOT EXISTS (SELECT * FROM sys.foreign_keys WHERE object_id = OBJECT_ID(N'[dbo].[FK__Usuarios__cod_ti__014935CB]') AND parent_object_id = OBJECT_ID(N'[dbo].[Usuarios]'))
ALTER TABLE [dbo].[Usuarios]  WITH CHECK ADD  CONSTRAINT [FK__Usuarios__cod_ti__014935CB] FOREIGN KEY([cod_tipo_usuario])
REFERENCES [dbo].[Tipo_Usuario] ([cod_tipo_usuario])
