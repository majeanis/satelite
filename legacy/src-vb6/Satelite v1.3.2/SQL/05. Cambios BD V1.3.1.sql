ALTER TABLE [dbo].[consultas]
ADD [num_area]          [integer] null
   ,[num_negocio]       [integer] null
   ,[ind_bloqueada]     [nvarchar] (01) COLLATE SQL_Latin1_General_CP1_CI_AS null
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tab_valores]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tab_valores]
GO

CREATE TABLE [dbo].[tab_valores]
    (
    [num_registro]   [integer] identity 
   ,[cod_tabla]      [nvarchar] (12) COLLATE SQL_Latin1_General_CP1_CI_AS null
   ,[gls_valor]      [nvarchar] (132) COLLATE SQL_Latin1_General_CP1_CI_AS null
    )
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[t_area]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[t_area]
GO

CREATE VIEW [dbo].[t_area] AS 
SELECT num_registro as num_area
      ,gls_valor    as gls_area
FROM   tab_valores
WHERE  cod_tabla = 'AREA'
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[t_negocio]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[t_negocio]
GO

CREATE VIEW [dbo].[t_negocio] AS 
SELECT num_registro as num_negocio
      ,gls_valor    as gls_negocio
FROM   tab_valores
WHERE  cod_tabla = 'NEGOCIO'
GO

----------------------------------------------------

SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER PROCEDURE [dbo].[usp_LeeConsultas] AS
DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
BEGIN
	--<V1.3.1>
	select c.*
	      ,b.nom_basedatos
	      ,isnull(a.gls_area,'') as gls_area
	      ,isnull(n.gls_negocio,'') as gls_negocio
	from   consultas c
	      ,basedatos b
	      ,t_area    a
	      ,t_negocio n
	where c.num_basedatos = b.num_basedatos
	and   a.num_area      =* c.num_area
	and   n.num_negocio   =* c.num_negocio
	--<\V1.3.1>
	order by c.num_consulta

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
	BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al seleccionar consultas'
		GOTO HandError
	END

	RETURN  
  
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
END
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_BloqueaConsulta]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE [dbo].[usp_BloqueaConsulta]
GO

CREATE PROCEDURE [dbo].[usp_BloqueaConsulta]
	(@wl_num_consulta	integer
	,@wl_ind_bloqueada	nvarchar(1)) AS
DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
BEGIN
	--<V1.3.1>
	UPDATE consultas
	SET ind_bloqueada  = @wl_ind_bloqueada
	WHERE num_consulta = @wl_num_consulta

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
	BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al bloquear consulta'
		GOTO HandError
	END

	RETURN  
  
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
	--<V1.3.1>
END
GO

ALTER PROCEDURE [dbo].[usp_ConsultasPorUsuario] 
	(@nom_usuario nvarchar(32)) AS
DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
BEGIN
	select c.num_consulta, c.nom_consulta, c.nom_dueno, cu.nom_creador, cu.fec_creacion
	      --<V1.3.1>
	      ,c.ind_bloqueada
	      --</V1.3.1>
	from cons_usuario cu
	    ,consultas    c
	where cu.nom_usuario = @nom_usuario
	and   c.num_consulta = cu.num_consulta
	order by c.nom_consulta

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
	BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al seleccionar consultas por usuario'
		GOTO HandError
	END

	RETURN  
  
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
END
GO

ALTER PROCEDURE [dbo].[usp_ConsultasPerfilPorUsuario]
    (@nom_usuario nvarchar(32) ) AS
DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
BEGIN
	select perfiles.num_perfil, perfiles.nom_perfil, consultas.num_consulta, consultas.nom_consulta
	       --<V1.3.1>
	      ,consultas.ind_bloqueada
	       --</V1.3.1>
	from perf_usuario
	    ,perfiles
	    ,cons_perfil
	    ,consultas
	where perf_usuario.nom_usuario = @nom_usuario
	and   perfiles.num_perfil      = perf_usuario.num_perfil
	and   cons_perfil.num_perfil   = perf_usuario.num_perfil
	and   consultas.num_consulta   = cons_perfil.num_consulta
	order by perfiles.nom_perfil, consultas.nom_consulta

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
	BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al seleccionar Perfil Por Usuario'
		GOTO HandError
	END

	RETURN  
  
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
END
GO

ALTER PROCEDURE [dbo].[usp_LeeConsultasPorLote]
	(@num_lote	integer
	) AS
DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
BEGIN
	--<V1.3.0>
	-- Selecciona todas las consultas del lote
	SELECT	C.*
	FROM	consultas c,
		cons_lote cl
	WHERE	cl.num_consulta = c.num_consulta 
	AND	cl.num_lote     = @num_lote
	--</V1.3.0>

	--<V1.3.1>
	AND    (c.ind_bloqueada is null or c.ind_bloqueada <> 'S')
	--</V1.3.1>


	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
	BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al seleccionar consultas por lote'
		GOTO HandError
	END

	RETURN  
  
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
END
GO

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

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_LeeTabValores]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE [dbo].[usp_LeeTabValores]
GO

CREATE PROCEDURE [dbo].[usp_LeeTabValores]
	(@wl_cod_tabla nvarchar(12)) AS
DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
BEGIN
	--<V1.3.1>
	IF @wl_cod_tabla = '' 
	BEGIN
		SELECT *
		FROM  tab_valores
		ORDER BY cod_tabla, gls_valor
	END
	ELSE
	BEGIN
		SELECT *
		FROM  tab_valores
		WHERE cod_tabla = @wl_cod_tabla
		ORDER BY gls_valor
	END

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
	BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al seleccionar tabla de valores'
		GOTO HandError
	END

	RETURN  
  
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
END
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_GrabaTabValores]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE [dbo].[usp_GrabaTabValores]
GO

CREATE PROCEDURE [dbo].[usp_GrabaTabValores]
	(@wl_num_registro nvarchar(12)
	,@wl_cod_tabla    nvarchar(12)
	,@wl_gls_valor    nvarchar(132)) AS
DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
BEGIN
	--<V1.3.1>
	IF (@wl_num_registro = '')
		INSERT INTO tab_valores
		    (cod_tabla
		    ,gls_valor
		    )
		VALUES
		    (@wl_cod_tabla
		    ,@wl_gls_valor 
		    )
	ELSE
		UPDATE tab_valores
		SET cod_tabla = @wl_cod_tabla
		   ,gls_valor = @wl_gls_valor
		WHERE num_registro = @wl_num_registro
	/* END IF */

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
	BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al insertar registro en tabla de valores'
		GOTO HandError
	END

	RETURN  
  
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
END
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_EliminaTabValores]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
DROP PROCEDURE [dbo].[usp_EliminaTabValores]
GO

CREATE PROCEDURE [dbo].[usp_EliminaTabValores]
	(@wl_num_registro integer) AS
DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
BEGIN
	--<V1.3.1>
	DELETE FROM tab_valores
	WHERE num_registro = @wl_num_registro

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
	BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar registro en tabla de valores'
		GOTO HandError
	END

	RETURN  
  
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
END
GO

-------------------------------------------------------------

UPDATE CONSULTAS
SET ind_bloqueada = 'N'
WHERE ind_bloqueada IS NULL

INSERT INTO tab_valores (cod_tabla, gls_valor) values ('ADMIN','AREA')
INSERT INTO tab_valores (cod_tabla, gls_valor) values ('ADMIN','NEGOCIO')

INSERT INTO tab_valores (cod_tabla, gls_valor) values ('AREA','OPERACIONES')
INSERT INTO tab_valores (cod_tabla, gls_valor) values ('AREA','COMERCIAL')
INSERT INTO tab_valores (cod_tabla, gls_valor) values ('AREA','TECNOLOGÍA')
INSERT INTO tab_valores (cod_tabla, gls_valor) values ('AREA','FINANZAS')
INSERT INTO tab_valores (cod_tabla, gls_valor) values ('AREA','INVERSIONES')

INSERT INTO tab_valores (cod_tabla, gls_valor) values ('NEGOCIO','VIDA INDIVIDUAL')
INSERT INTO tab_valores (cod_tabla, gls_valor) values ('NEGOCIO','COLECTIVOS')
INSERT INTO tab_valores (cod_tabla, gls_valor) values ('NEGOCIO','BANCA SEGURO')
INSERT INTO tab_valores (cod_tabla, gls_valor) values ('NEGOCIO','SOAT')
INSERT INTO tab_valores (cod_tabla, gls_valor) values ('NEGOCIO','RENTAS VITALICIAS')
INSERT INTO tab_valores (cod_tabla, gls_valor) values ('NEGOCIO','GENERALES')
