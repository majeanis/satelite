ALTER TABLE [dbo].[consultas]
ADD gls_horario_ejecucion nvarchar(40) COLLATE SQL_Latin1_General_CP1_CI_AS null
GO

ALTER PROCEDURE [dbo].[usp_GrabaConsulta]
	(@num_consulta			integer
	,@nom_consulta			nvarchar(132)
	,@num_basedatos			integer
	,@gls_query				ntext
	,@gls_parametros			ntext
	,@gls_formatos			ntext
	,@gls_horario_ejecucion		nvarchar(40)
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
	SELECT	 @num_consulta
		,nom_parametro
		,cod_tipo_dato
		,cod_tipo_ayuda
		,gls_ayuda_valores
		,ind_opcional
	FROM	OPENXML (@idoc, '/ROOT/Parametros',1)
	WITH	(num_consulta		integer
		,nom_parametro		nvarchar(80)
		,cod_tipo_dato		nvarchar(12)
		,cod_tipo_ayuda		nvarchar(12)
		,gls_ayuda_valores	ntext
		,ind_opcional		nvarchar(1))

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
