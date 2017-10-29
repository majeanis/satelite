if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Cons_Lote_Lotes]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Cons_Lote] DROP CONSTRAINT FK_Cons_Lote_Lotes
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Lote_Usuario_Lotes]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[Lote_Usuario] DROP CONSTRAINT FK_Lote_Usuario_Lotes
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_GrabaLote]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_GrabaLote]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_LeeTodasConsultasPorLote]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_LeeTodasConsultasPorLote]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_GrabaUsuariosPorLote]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_GrabaUsuariosPorLote]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_EliminaUsuariosPorLote]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_EliminaUsuariosPorLote]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_UsuariosPorLote]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_UsuariosPorLote]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_UsuariosSinLote]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_UsuariosSinLote]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_EliminaLotesPorUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_EliminaLotesPorUsuario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_LotesPorUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_LotesPorUsuario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_LotesSinUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_LotesSinUsuario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_LeeLotes]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_LeeLotes]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_EliminaLote]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_EliminaLote]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cons_Lote]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Cons_Lote]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Lote_Usuario]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Lote_Usuario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Lotes]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Lotes]
GO

CREATE TABLE [dbo].[Cons_Lote] (
	[num_lote] [int] NOT NULL ,
	[num_consulta] [int] NOT NULL ,
	[fec_creacion] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Lote_Usuario] (
	[nom_usuario] [nvarchar] (32) COLLATE Modern_Spanish_CI_AI NOT NULL ,
	[num_lote] [int] NOT NULL ,
	[fec_creacion] [datetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Lotes] (
	[num_lote] [int] IDENTITY (1, 1) NOT NULL ,
	[nom_lote] [nvarchar] (132) COLLATE Modern_Spanish_CI_AI NULL ,
	[nom_creador] [nvarchar] (32) COLLATE Modern_Spanish_CI_AI NULL ,
	[nom_solicitante] [nvarchar] (32) COLLATE Modern_Spanish_CI_AI NULL ,
	[fec_creacion] [datetime] NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Cons_Lote] WITH NOCHECK ADD 
	CONSTRAINT [PK_Cons_Grupo] PRIMARY KEY  CLUSTERED 
	(
		[num_lote],
		[num_consulta]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Lote_Usuario] WITH NOCHECK ADD 
	CONSTRAINT [PK_Grup_Usuario] PRIMARY KEY  CLUSTERED 
	(
		[nom_usuario],
		[num_lote]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[Lotes] WITH NOCHECK ADD 
	CONSTRAINT [PK_Lotes] PRIMARY KEY  CLUSTERED 
	(
		[num_lote]
	)  ON [PRIMARY] 
GO

 CREATE  INDEX [IX_Cons_Lote] ON [dbo].[Cons_Lote]([num_consulta]) ON [PRIMARY]
GO

 CREATE  INDEX [IX_Lote_Usuario] ON [dbo].[Lote_Usuario]([num_lote]) ON [PRIMARY]
GO

ALTER TABLE [dbo].[Cons_Lote] ADD 
	CONSTRAINT [FK_Cons_Lote_Consultas] FOREIGN KEY 
	(
		[num_consulta]
	) REFERENCES [dbo].[Consultas] (
		[num_consulta]
	),
	CONSTRAINT [FK_Cons_Lote_Lotes] FOREIGN KEY 
	(
		[num_lote]
	) REFERENCES [dbo].[Lotes] (
		[num_lote]
	)
GO

ALTER TABLE [dbo].[Lote_Usuario] ADD 
	CONSTRAINT [FK_Lote_Usuario_Lotes] FOREIGN KEY 
	(
		[num_lote]
	) REFERENCES [dbo].[Lotes] (
		[num_lote]
	)
GO

ALTER TABLE [dbo].[consultas]
ADD [gls_archivo_salida]     [nvarchar] (500) null
   ,[nom_hoja_salida]        [nvarchar] (40) null
GO

-------------------------------------------------------------

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

ALTER PROCEDURE usp_GrabaConsulta
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
		   --<V1.3.0>
		   ,gls_archivo_salida      = @gls_archivo_salida
		   ,nom_hoja_salida         = @nom_hoja_salida
		   --</V1.3.0>
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
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_GrabaLote]
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
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al validar nombre del lote'
			GOTO HandError
			END
	
		IF @wl_num_filas > 0
			BEGIN
			SELECT @wl_gls_error = 'Nombre del lote ya existe. Intente con otro nombre'
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
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al insertar lote'
			GOTO HandError
			END

		-- Obtiene el numero de lote asignado
		SELECT @num_lote = num_lote
		FROM   Lotes
		WHERE  nom_lote = @nom_lote
	
		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al obtener número de lote'
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
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al actualizar lote'
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
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar consultas del lote'
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
	FROM OPENXML (@idoc, '/ROOT/ConsLote',1)
	WITH (num_lote     integer
	     ,num_consulta integer)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al crear consultas del lote'
		GOTO HandError
		END

	-- Asigna al usuario solicitante al lote, en caso que se indique
	IF @wl_ind_asignar_lote = 'S'
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
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al asignar usuario solicitante al lote'
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_LeeTodasConsultasPorLote]
	(@num_lote	integer
	) AS
BEGIN
	--<V1.3.0>
	-- Selecciona todas las consultas y aquellas que ya estan asociadas al lote las marca con una N

	SELECT	C.*
	       ,CASE WHEN ISNULL(cl.num_lote,0)=0 THEN 'N' ELSE 'S' END as ind_lote
	FROM	consultas c,
		cons_lote cl
	WHERE	cl.num_consulta =* c.num_consulta 
	AND	cl.num_lote     = @num_lote

	--</V1.3.0>
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_LeeConsultasPorLote]
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_GrabaUsuariosPorLote]
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
	FROM OPENXML (@idoc, '/ROOT/LoteUsuario', 1)  
	WITH (num_lote     integer  
	     ,nom_usuario  nvarchar(32)  
	     )  
	
	SET @wl_num_error = @@ERROR  
	IF @wl_num_error <> 0  
		BEGIN  
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al insertar usuarios por lote'  
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_EliminaUsuariosPorLote]
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
	      FROM OPENXML (@idoc, '/ROOT/LoteUsuario',1)  
	      WITH (num_lote    integer  
	           ,nom_usuario nvarchar(32))  
	     )  
	
	SET @wl_num_error = @@ERROR  
	IF @wl_num_error <> 0  
		BEGIN  
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar usuarios por lote'  
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_UsuariosPorLote]
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
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al consultar Usuarios por Lote'  
		GOTO HandError  
	END  

	RETURN  
	
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
	--</V1.3.0>
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_UsuariosSinLote]
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
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al consultar Usuarios no asignados al Lote'  
		GOTO HandError  
	END  

	RETURN  
	
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
	--</V1.3.0>
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_EliminaLotesPorUsuario]
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
	      FROM OPENXML (@idoc, '/ROOT/LoteUsuario',1)  
	      WITH (num_lote    integer  
	           ,nom_usuario nvarchar(32))  
	     )  
	
	SET @wl_num_error = @@ERROR  
	IF @wl_num_error <> 0  
		BEGIN  
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar lotes por usuario'
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_LotesPorUsuario]
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
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al consultar Lotes por Usuarios'  
		GOTO HandError  
	END  

	RETURN  
	
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
	--</V1.3.0>
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


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
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al consultar Lotes no asignados al Usuario'  
		GOTO HandError  
	END  

	RETURN  
	
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
	--</V1.3.0>
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_LeeLotes] AS
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
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al consultar Lotes'
		GOTO HandError  
	END  

	RETURN  
	
HandError:  
	ROLLBACK TRANSACTION  
	RAISERROR(@wl_gls_error,16,1)  
	RETURN  
	--</V1.3.0>
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_EliminaLote]
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
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar consultas del lote'
		GOTO HandError
		END

	-- Elimina lote
	DELETE FROM lotes
	WHERE  num_lote = @num_lote

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar lote'
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

