ALTER TABLE [dbo].[parametros]
ADD gls_parametro nvarchar(132) COLLATE SQL_Latin1_General_CP1_CI_AS NULL
GO

CREATE TABLE [dbo].[Carpetas](
	[num_carpeta] [int] IDENTITY(1,1) NOT NULL,
	[nom_usuario] [nvarchar](32) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	[gls_carpeta] [nvarchar](300) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Cons_Carpeta](
	[nom_usuario] [nvarchar](32) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	[num_consulta] [int] NOT NULL,
	[num_carpeta] [int] NULL
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Carpetas] WITH NOCHECK ADD 
	CONSTRAINT [PK_Carpetas] PRIMARY KEY  CLUSTERED 
	(
		[num_carpeta] 
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Cons_Carpeta] WITH NOCHECK ADD 
 CONSTRAINT [PK_Cons_Carpeta] PRIMARY KEY CLUSTERED 
(
	[nom_usuario] ,
	[num_consulta] 
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Cons_Carpeta]  ADD  
	CONSTRAINT [FK_Cons_Carpeta_Carpetas] FOREIGN KEY
	([num_carpeta]
	) REFERENCES [dbo].[Carpetas] ([num_carpeta])
GO

ALTER TABLE [dbo].[Carpetas]  ADD  
	CONSTRAINT [FK_Carpetas_Usuarios] FOREIGN KEY
	([nom_usuario]
	) REFERENCES [dbo].[Usuarios] ([nom_usuario])
GO

CREATE PROCEDURE [dbo].[usp_GrabaCarpetaUsuario]
	(@num_carpeta		integer
	,@nom_usuario		nvarchar(32)
	,@gls_carpeta		nvarchar(300)
	) as

DECLARE
	 @wl_num_error		integer 
	,@wl_gls_error		nvarchar(132)
	,@wl_num_filas		integer

BEGIN
--<V1.3.2>
	SET NOCOUNT ON
	BEGIN TRANSACTION
	
	-- Valida que el nombre del tipo de usuario sea único
	IF @num_carpeta = 0
		BEGIN

		SELECT @wl_num_filas = COUNT(*)
		FROM   carpetas
		WHERE  nom_usuario = @nom_usuario
		AND    gls_carpeta = @gls_carpeta
	
		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al validar nombre de la carpeta'
			GOTO HandError
			END
	
		IF @wl_num_filas > 0
			BEGIN
			SELECT @wl_gls_error = 'Nombre de la carpeta ya existe. Intente con otro nombre'
			GOTO HandError
			END

		-- Crea usuario
		INSERT INTO carpetas
		       (nom_usuario 
		       ,gls_carpeta 
		       )
		VALUES (@nom_usuario 
		       ,@gls_carpeta
		       )

		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al insertar carpeta'
			GOTO HandError
			END
		END
	ELSE
		BEGIN

		UPDATE carpetas
		SET	 nom_usuario = @nom_usuario 
			,gls_carpeta = @gls_carpeta
		WHERE  num_carpeta = @num_carpeta

		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al actualizar la carpeta'
			GOTO HandError
			END

		END
	/* END IF */

	COMMIT TRANSACTION
	SET NOCOUNT OFF

	-- Devuelve la información de la consulta creada o actualizada
	SELECT *
	FROM   carpetas
	WHERE  nom_usuario = @nom_usuario
	AND    gls_carpeta = @gls_carpeta

	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
--</V1.3.2>
END

GO

CREATE PROCEDURE [dbo].[usp_EliminaCarpetaUsuario]
	(@nom_usuario nvarchar(32)
	,@gls_carpeta nvarchar(300)) as
DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
BEGIN
--<V1.3.2>
	SET NOCOUNT ON
	BEGIN TRANSACTION

	-- Elimina carpeta del usuario
	DELETE FROM carpetas
	WHERE  nom_usuario = @nom_usuario
	AND    gls_carpeta = @gls_carpeta

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar carpeta del usuario'
		GOTO HandError
		END

	-- Elimina carpetas anidadas dentro de la carpeta del usuario
	DELETE FROM carpetas
	WHERE  nom_usuario = @nom_usuario
	AND    gls_carpeta like @gls_carpeta + '\%'

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar carpetas anidadas'
		GOTO HandError
		END

	COMMIT TRANSACTION
	SET NOCOUNT OFF

	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
--</V1.3.2>
END

GO

CREATE PROCEDURE [dbo].[usp_LeeCarpetasUsuario] 
	(@nom_usuario nvarchar(32)
	) AS
DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
BEGIN
--<V1.3.2>
	SELECT *
	FROM   carpetas
	WHERE  nom_usuario = @nom_usuario
	ORDER BY gls_carpeta

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
	BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al seleccionar carpetas'
		GOTO HandError
	END

	RETURN  
  
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
--</V1.3.2>
END

GO

---------------------------------------------------------------------

CREATE PROCEDURE [dbo].[usp_GrabaConsultaCarpeta]
	(@nom_usuario		nvarchar(32)
	,@num_consulta		integer
	,@num_carpeta		integer
	) AS

DECLARE
	 @wl_num_error		integer 
	,@wl_gls_error		nvarchar(132)
	,@wl_num_filas		integer

BEGIN
--<V1.3.2>
	SET NOCOUNT ON
	BEGIN TRANSACTION

	-- Valida si existe la consulta en una carpeta
	SELECT @wl_num_filas = COUNT(*)
	FROM   cons_carpeta
	WHERE  nom_usuario  = @nom_usuario
	AND    num_consulta = @num_consulta

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al validar consulta-carpeta'
		GOTO HandError
		END
	/* END IF */
	
	IF @wl_num_filas = 0
		BEGIN
		-- Crea consulta carpeta
		INSERT INTO cons_carpeta
		       (nom_usuario 
		       ,num_consulta
		       ,num_carpeta 
		       )
		VALUES (@nom_usuario 
		       ,@num_consulta
		       ,@num_carpeta
		       )

		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al insertar consulta-carpeta'
			GOTO HandError
			END
		END
	ELSE
		BEGIN

		UPDATE cons_carpeta
		SET	 num_carpeta = @num_carpeta
		WHERE  nom_usuario  = @nom_usuario 
		AND    num_consulta = @num_consulta

		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al actualizar la consulta-carpeta'
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
--</V1.3.2>
END

GO


CREATE PROCEDURE [dbo].[usp_LeeConsultasEnCarpetas]
	(@nom_usuario		nvarchar(32)
	) AS

DECLARE
	 @wl_num_error		integer 
	,@wl_gls_error		nvarchar(132)
	,@wl_num_filas		integer

BEGIN
--<V1.3.2>
	SELECT ca.nom_usuario, ca.num_carpeta, ca.gls_carpeta, cc.num_consulta
	FROM   carpetas     ca
	      ,cons_carpeta cc
	WHERE  ca.nom_usuario  = @nom_usuario
	AND    cc.nom_usuario  = @nom_usuario
	AND    cc.num_carpeta  = ca.num_carpeta

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al obtener consultas en carpeta'
		GOTO HandError
		END
	/* END IF */	
	
	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
--</V1.3.2>
END

GO

CREATE PROCEDURE [dbo].[usp_EliminaConsultaEnCarpeta]
	(@nom_usuario  nvarchar(32)
	,@num_consulta integer) AS
DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
BEGIN
--<V1.3.2>
	SET NOCOUNT ON
	BEGIN TRANSACTION

	-- Elimina la consulta en la carpeta del usuario
	DELETE FROM cons_carpeta
	WHERE  nom_usuario  = @nom_usuario
	AND    num_consulta = @num_consulta

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar consulta de la carpeta del usuario'
		GOTO HandError
		END

	COMMIT TRANSACTION
	SET NOCOUNT OFF

	RETURN

HandError:
	ROLLBACK TRANSACTION
	RAISERROR(@wl_gls_error,16,1)
	RETURN
--</V1.3.2>
END

GO

CREATE TRIGGER tg_del_carpetas ON dbo.carpetas
	AFTER DELETE
AS 
BEGIN
--<V1.3.2>
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

	-- Marca las carpetas eliminadas
	UPDATE Cons_Carpeta
	SET num_carpeta = -1
	FROM deleted
	WHERE Cons_Carpeta.nom_usuario = deleted.nom_usuario
	AND   Cons_Carpeta.num_carpeta = deleted.num_carpeta

	-- Elimina las carpetas
	DELETE FROM Cons_Carpeta
	WHERE num_carpeta = -1
--</V1.3.2>
END

GO

--------------------------------------------

ALTER PROCEDURE [dbo].[usp_GrabaConsulta]
	(@num_consulta			integer
	,@nom_consulta			nvarchar(132)
	,@num_basedatos			integer
	,@gls_query				ntext
	,@gls_parametros		ntext
	,@gls_formatos			ntext
	,@gls_horario_ejecucion	nvarchar(40)
--<V1.3.0>
	,@gls_archivo_salida	nvarchar(500)
	,@nom_hoja_salida		nvarchar(40)
--</V1.3.0>
--<V1.3.1>
	,@num_area				integer
	,@num_negocio			integer
--</V1.3.1>
	,@nom_user				nvarchar(32)
	,@nom_user_real			nvarchar(32)
	) AS

DECLARE
	 @wl_ind_asignar_consulta	char(1)
	,@wl_num_error				integer 
	,@wl_gls_error				nvarchar(132)
	,@wl_num_filas				integer
	,@idoc						integer

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
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al validar nombre de la consulta'
			GOTO HandError
			END
	
		IF @wl_num_filas > 0
			BEGIN
			SELECT @wl_gls_error = 'Nombre de la consulta ya existe. Intente con otro nombre'
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
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al obtener detalle de usuario'
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
			--<V1.3.1>
			,num_area
			,num_negocio
			,ind_bloqueada
			--</V1.3.1>
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
			--<V1.3.1>
			,@num_area
			,@num_negocio
			,'N'
			--</V1.3.1>
			)

		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al insertar consulta'
			GOTO HandError
			END

		-- Obtiene el numero de consulta asignado
		SELECT @num_consulta = num_consulta
		FROM   consultas
		WHERE  nom_consulta = @nom_consulta
	
		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al obtener número de consulta'
			GOTO HandError
			END

		-- Si el perfil, asigna automáticamente la cosulta al dueño, ingresa la consulta al usuario
		IF @wl_ind_asignar_consulta = 'S'
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
				SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al obtener detalle de usuario'
				GOTO HandError
				END
		/* END if */
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
		   --<V1.3.1>
		   ,num_area                = @num_area
		   ,num_negocio             = @num_negocio
		   --</V1.3.1>
		WHERE num_consulta = @num_consulta

		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al actualizar consulta'
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
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar parametros'
		GOTO HandError
		END

	INSERT INTO parametros
		(num_consulta
		,nom_parametro
		,cod_tipo_dato
		,cod_tipo_ayuda
		,gls_ayuda_valores
		,ind_opcional
--<V1.3.2>
		,gls_parametro
--</V1.3.2>
		)
	SELECT	 @num_consulta
		,nom_parametro
		,cod_tipo_dato
		,cod_tipo_ayuda
		,gls_ayuda_valores
		,ind_opcional
--<V1.3.2>
		,gls_parametro
--</V1.3.2>
	FROM	OPENXML (@idoc, '/ROOT/Parametros',1)
	WITH	(num_consulta		integer
		,nom_parametro		nvarchar(80)
		,cod_tipo_dato		nvarchar(12)
		,cod_tipo_ayuda		nvarchar(12)
		,gls_ayuda_valores	ntext
		,ind_opcional		nvarchar(1)
--<V1.3.2>
		,gls_parametro		nvarchar(132))
--</V1.3.2>

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al crear parámetros'
		GOTO HandError
		END

	-- Graba Formatos de la consulta
	EXEC usp_GrabaFormatosConsulta @num_consulta, @gls_formatos

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al grabar formatos'
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

GO

CREATE PROCEDURE [dbo].[usp_LeeConsultaPorNombre]
	(@nom_consulta nvarchar(132)) as
BEGIN
--<V1.3.2>
	SELECT *
	FROM  consultas   
	WHERE nom_consulta = @nom_consulta
--</V1.3.2>
END

GO

