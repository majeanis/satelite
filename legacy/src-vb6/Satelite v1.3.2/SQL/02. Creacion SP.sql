if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_ConsultasPerfilPorUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_ConsultasPerfilPorUsuario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_ConsultasPorPerfil]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_ConsultasPorPerfil]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_ConsultasPorUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_ConsultasPorUsuario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_ConsultasSinPerfil]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_ConsultasSinPerfil]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_ConsultasSinUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_ConsultasSinUsuario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_DetalleConsulta]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_DetalleConsulta]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_DetalleParametro]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_DetalleParametro]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_DetalleUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_DetalleUsuario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_EliminaBaseDatos]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_EliminaBaseDatos]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_EliminaConsulta]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_EliminaConsulta]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_EliminaConsultasPorPerfil]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_EliminaConsultasPorPerfil]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_EliminaConsultasPorUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_EliminaConsultasPorUsuario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_EliminaPerfil]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_EliminaPerfil]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_EliminaPerfilesPorConsulta]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_EliminaPerfilesPorConsulta]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_EliminaPerfilesPorUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_EliminaPerfilesPorUsuario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_EliminaTipoUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_EliminaTipoUsuario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_EliminaUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_EliminaUsuario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_EliminaUsuariosPorConsulta]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_EliminaUsuariosPorConsulta]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_EliminaUsuariosPorPerfil]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_EliminaUsuariosPorPerfil]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_GrabaBaseDatos]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_GrabaBaseDatos]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_GrabaConsulta]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_GrabaConsulta]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_GrabaConsultasPorPerfiles]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_GrabaConsultasPorPerfiles]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_GrabaConsultasPorUsuarios]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_GrabaConsultasPorUsuarios]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_GrabaFormatosConsulta]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_GrabaFormatosConsulta]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_GrabaLogEjecucion]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_GrabaLogEjecucion]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_GrabaPerfil]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_GrabaPerfil]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_GrabaPerfilesPorUsuarios]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_GrabaPerfilesPorUsuarios]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_GrabaTipoUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_GrabaTipoUsuario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_GrabaUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_GrabaUsuario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_LeeBaseDatos]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_LeeBaseDatos]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_LeeBasesDeDatos]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_LeeBasesDeDatos]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_LeeConsultas]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_LeeConsultas]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_LeeFormatos]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_LeeFormatos]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_LeePerfiles]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_LeePerfiles]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_LeeProveedores]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_LeeProveedores]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_LeeSysConfig]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_LeeSysConfig]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_LeeTipoDatos]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_LeeTipoDatos]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_LeeTipoUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_LeeTipoUsuario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_LeeTiposUsuarios]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_LeeTiposUsuarios]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_LeeUsuarios]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_LeeUsuarios]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_PerfilesPorConsulta]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_PerfilesPorConsulta]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_PerfilesPorUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_PerfilesPorUsuario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_PerfilesSinConsulta]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_PerfilesSinConsulta]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_PerfilesSinUsuario]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_PerfilesSinUsuario]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_UsuariosPorConsulta]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_UsuariosPorConsulta]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_UsuariosPorPerfil]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_UsuariosPorPerfil]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_UsuariosSinConsulta]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_UsuariosSinConsulta]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[usp_UsuariosSinPerfil]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[usp_UsuariosSinPerfil]
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_ConsultasPerfilPorUsuario]
	(@nom_usuario nvarchar(32) ) AS
BEGIN
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

CREATE PROCEDURE [dbo].[usp_ConsultasPorPerfil]
	(@num_perfil integer) AS
BEGIN
	select c.num_consulta, c.nom_consulta, cp.fec_creacion
	from perfiles    p
	    ,cons_perfil cp
	    ,consultas   c
	where p.num_perfil   = @num_perfil
	and   cp.num_perfil  = p.num_perfil
	and   c.num_consulta = cp.num_consulta
	order by c.nom_consulta
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

CREATE PROCEDURE [dbo].[usp_ConsultasPorUsuario]
	(@nom_usuario nvarchar(32)) AS
BEGIN
	select c.num_consulta, c.nom_consulta, c.nom_dueno, cu.nom_creador, cu.fec_creacion
	from cons_usuario cu
	    ,consultas    c
	where cu.nom_usuario = @nom_usuario
	and   c.num_consulta = cu.num_consulta
	order by c.nom_consulta
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

CREATE PROCEDURE [dbo].[usp_ConsultasSinPerfil]
	(@num_perfil integer) AS
BEGIN
	select *
	from consultas
	where num_consulta not in (
	select distinct num_consulta
	from cons_perfil
	where num_perfil   = @num_perfil
	)
	order by nom_consulta
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

CREATE PROCEDURE [dbo].[usp_ConsultasSinUsuario]
	(@nom_usuario nvarchar(32)) AS
BEGIN
	select *
	from consultas
	where num_consulta not in (
	select distinct num_consulta
	from cons_usuario
	where nom_usuario = @nom_usuario
	)
	order by nom_consulta
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

CREATE PROCEDURE [dbo].[usp_DetalleConsulta]
	(@num_consulta integer) AS
BEGIN
	select *
	from   consultas   
	where  num_consulta = @num_consulta
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

CREATE PROCEDURE [dbo].[usp_DetalleParametro]
	(@num_consulta	integer
	,@nom_parametro nvarchar(32)) AS
BEGIN
	select *
	from   parametros
	where  num_consulta  = @num_consulta
	and    nom_parametro = @nom_parametro
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

CREATE PROCEDURE [dbo].[usp_DetalleUsuario]
	(@nom_usuario nvarchar(32) ) AS
BEGIN
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

CREATE PROCEDURE [dbo].[usp_EliminaBaseDatos]
	(@num_basedatos integer) AS
DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
BEGIN
	SET NOCOUNT ON
	BEGIN TRANSACTION

	-- Elimina base de datos
	DELETE FROM BaseDatos
	WHERE  num_basedatos = @num_basedatos

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar base de datos'
		GOTO HandError
		END

	COMMIT TRANSACTION
	SET NOCOUNT OFF

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

CREATE PROCEDURE [dbo].[usp_EliminaConsulta]
	(@num_consulta integer) AS
DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
BEGIN
	SET NOCOUNT ON
	BEGIN TRANSACTION

	-- Elimina parámetros de la consulta
	DELETE FROM parametros
	WHERE  num_consulta = @num_consulta

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar parámetros de la consulta'
		GOTO HandError
		END

	-- Elimina la consulta
	DELETE FROM consultas
	WHERE  num_consulta = @num_consulta

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar la consulta'
		GOTO HandError
		END

	COMMIT TRANSACTION
	SET NOCOUNT OFF

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

CREATE PROCEDURE [dbo].[usp_EliminaConsultasPorPerfil]
	(@num_perfil	integer
	,@gls_consultas	ntext
	) AS

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
		 FROM	OPENXML (@idoc, '/ROOT/ConsPerfil',1)
		 WITH	(num_perfil		integer
			,num_consulta		integer)
		)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar consultas por perfil'
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

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
		 FROM	OPENXML (@idoc, '/ROOT/ConsUsuario',1)
		 WITH	(num_consulta		integer
			,nom_usuario		nvarchar(32))
		)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar consultas por usuario'
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_EliminaPerfil]
	(@num_perfil integer) AS
DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
BEGIN
	SET NOCOUNT ON
	BEGIN TRANSACTION

	-- Elimina perfil
	DELETE FROM perfiles
	WHERE  num_perfil = @num_perfil

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar perfil'
		GOTO HandError
		END

	COMMIT TRANSACTION
	SET NOCOUNT OFF

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

CREATE PROCEDURE [dbo].[usp_EliminaPerfilesPorConsulta]
	(@num_consulta	integer
	,@gls_perfiles	ntext
	) AS

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
		 FROM	OPENXML (@idoc, '/ROOT/ConsPerfil',1)
		 WITH	(num_perfil		integer
			,num_consulta		integer)
		)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar perfiles por consulta'
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_EliminaPerfilesPorUsuario]
	(@nom_usuario	nvarchar(32)
	,@gls_perfiles	ntext
	) AS

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
		 FROM	OPENXML (@idoc, '/ROOT/PerfUsuario',1)
		 WITH	(num_perfil		integer
			,nom_usuario		nvarchar(32))
		)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar perfiles por usuario'
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_EliminaTipoUsuario]
	(@cod_tipo_usuario nvarchar(12)) AS
DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
BEGIN
	SET NOCOUNT ON
	BEGIN TRANSACTION

	-- Elimina tipo usuario
	DELETE FROM tipo_usuario
	WHERE  cod_tipo_usuario = @cod_tipo_usuario

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar tipo usuario'
		GOTO HandError
		END

	COMMIT TRANSACTION
	SET NOCOUNT OFF

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

CREATE PROCEDURE [dbo].[usp_EliminaUsuario]
	(@nom_usuario nvarchar(32)) AS
DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
BEGIN
	SET NOCOUNT ON
	BEGIN TRANSACTION

	-- Elimina usuario
	DELETE FROM usuarios
	WHERE  nom_usuario = @nom_usuario

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar usuario'
		GOTO HandError
		END

	COMMIT TRANSACTION
	SET NOCOUNT OFF

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

CREATE PROCEDURE [dbo].[usp_EliminaUsuariosPorConsulta]
	(@num_consulta	integer
	,@gls_usuarios	ntext
	) AS

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
		 FROM	OPENXML (@idoc, '/ROOT/ConsUsuario',1)
		 WITH	(num_consulta		integer
			,nom_usuario		nvarchar(32))
		)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar usuarios por consulta'
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_EliminaUsuariosPorPerfil]
	(@num_perfil	integer
	,@gls_usuarios	ntext
	) AS

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
		 FROM	OPENXML (@idoc, '/ROOT/PerfUsuario',1)
		 WITH	(num_perfil		integer
			,nom_usuario		nvarchar(32))
		)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar usuarios por perfil'
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE  PROCEDURE [dbo].[usp_GrabaBaseDatos]
	(@num_basedatos		integer
	,@nom_basedatos		nvarchar(80)
	,@gls_coneccion		nvarchar(500)
	,@gls_formato_fecha	nvarchar(32)
	) AS

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
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al validar nombre de la base de datos'
			GOTO HandError
			END
	
		IF @wl_num_filas > 0
			BEGIN
			SELECT @wl_gls_error = 'Nombre de la base de datos ya existe. Intente con otro nombre'
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
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al insertar base de datos'
			GOTO HandError
			END

		-- Obtiene el numero de la base de datos asignado
		SELECT @num_basedatos = num_basedatos
		FROM   basedatos
		WHERE  nom_basedatos = @nom_basedatos
	
		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al obtener número de la base de datos'
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
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al validar nombre de la base de datos'
			GOTO HandError
			END
	
		IF @wl_num_filas > 0
			BEGIN
			SELECT @wl_gls_error = 'Nombre de la base de datos ya existe. Intente con otro nombre'
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
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al actualizar base de datos'
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_GrabaFormatosConsulta]
	(@num_consulta	integer
	,@gls_formatos		ntext
	) AS

DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
	,@wl_num_filas			integer
	,@idoc					integer

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
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al eliminar formato de columnas'
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
	FROM	OPENXML (@idoc, '/ROOT/Formatos',1)
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
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al insertar formato de columnas'
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_GrabaConsulta]
	(@num_consulta		integer
	,@nom_consulta		nvarchar(132)
	,@num_basedatos		integer
	,@gls_query			ntext
	,@gls_parametros		ntext
	,@gls_formatos		ntext
	,@nom_user			nvarchar(32)
	,@nom_user_real		nvarchar(32)
	) AS

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
			)
		VALUES
			(@nom_consulta
			,@num_basedatos
			,@gls_query
			,@nom_user
			,@nom_user_real
			,getdate()
			,getdate()
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
		SET nom_consulta          = @nom_consulta
		   ,num_basedatos         = @num_basedatos
		   ,gls_query             = @gls_query
		   ,fec_ult_actualizacion = getdate()
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
	FROM	OPENXML (@idoc, '/ROOT/ConsPerfil',1)
	WITH	(num_perfil		integer
		,num_consulta		integer
		)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al insertar consultas por perfiles'
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_GrabaConsultasPorUsuarios]
	(@nom_user			nvarchar(32)
	,@gls_usuarios		ntext
	) AS

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
	FROM	OPENXML (@idoc, '/ROOT/ConsUsuario',1)
	WITH	(num_consulta		integer
		,nom_usuario		nvarchar(32)
		)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al insertar consultas por usuarios'
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_GrabaLogEjecucion]
	(@num_consulta		integer
	,@nom_usuario		nvarchar(32)
	,@fec_ejecucion		nvarchar(12)
	,@hor_inicio		nvarchar(12)
	,@gls_tiempo_utilizado	nvarchar(12)
	,@num_registros		integer
	) AS

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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_GrabaPerfil]
	(@num_perfil		integer
	,@nom_perfil		nvarchar(132)
	) AS

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
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al validar nombre del perfil'
			GOTO HandError
			END
	
		IF @wl_num_filas > 0
			BEGIN
			SELECT @wl_gls_error = 'Nombre del perfil ya existe. Intente con otro nombre'
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
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al insertar perfil'
			GOTO HandError
			END

		-- Obtiene el numero del perfil asignado
		SELECT @num_perfil = num_perfil
		FROM   perfiles
		WHERE  nom_perfil = @nom_perfil
	
		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al obtener número del perfil'
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
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al actualizar perfil'
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_GrabaPerfilesPorUsuarios]
	(@nom_user			nvarchar(32)
	,@gls_usuarios		ntext
	) AS

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
	FROM	OPENXML (@idoc, '/ROOT/PerfUsuario',1)
	WITH	(num_perfil		integer
		,nom_usuario		nvarchar(32)
		)

	SET @wl_num_error = @@ERROR
	IF @wl_num_error <> 0
		BEGIN
		SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al insertar perfiles por usuarios'
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_GrabaTipoUsuario]
	(@cod_tipo_usuario	nvarchar(12)
	,@gls_campos		ntext
	,@cod_tipo_accion	nvarchar(08) 
	) AS

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
	IF @cod_tipo_accion = 'INS'
		BEGIN

		SELECT @wl_num_filas = COUNT(*)
		FROM   tipo_usuario
		WHERE  cod_tipo_usuario = @cod_tipo_usuario
	
		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al validar nombre del tipo de usuario'
			GOTO HandError
			END
	
		IF @wl_num_filas > 0
			BEGIN
			SELECT @wl_gls_error = 'Nombre del tipo de usuario ya existe. Intente con otro nombre'
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
		FROM	OPENXML (@idoc, '/ROOT/TipoUsuario',1)
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
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al insertar tipo de usuario'
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
		FROM	OPENXML (@idoc, '/ROOT/TipoUsuario',1)
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
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al actualizar tipo de usuario'
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_GrabaUsuario]
	(@nom_usuario		nvarchar(32)
	,@cod_tipo_usuario	nvarchar(32)
	,@cod_tipo_accion	nvarchar(08) 
	) AS

DECLARE
	 @wl_num_error			integer 
	,@wl_gls_error			nvarchar(132)
	,@wl_num_filas			integer

BEGIN
	SET NOCOUNT ON
	BEGIN TRANSACTION
	
	-- Valida que el nombre de la consulta sea único (aun cuando lo hace el indice, se valida para mejorar el mensaje de error)
	IF @cod_tipo_accion = 'INS'
		BEGIN

		SELECT @wl_num_filas = COUNT(*)
		FROM   usuarios
		WHERE  nom_usuario = @nom_usuario
	
		SET @wl_num_error = @@ERROR
		IF @wl_num_error <> 0
			BEGIN
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al validar nombre del usuario'
			GOTO HandError
			END
	
		IF @wl_num_filas > 0
			BEGIN
			SELECT @wl_gls_error = 'Nombre del usuario ya existe. Intente con otro nombre'
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
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al insertar usuario'
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
			SELECT @wl_gls_error = CAST(@@ERROR AS NVARCHAR(10)) + ', error al actualizar usuario'
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

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO

CREATE PROCEDURE [dbo].[usp_LeeBaseDatos]
	(@num_basedatos	integer) AS
BEGIN
	select *
	from   BaseDatos
	where  num_basedatos = @num_basedatos
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

CREATE PROCEDURE [dbo].[usp_LeeBasesDeDatos] AS
BEGIN
	select *
	from   BaseDatos
	order by num_basedatos
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

CREATE PROCEDURE [dbo].[usp_LeeConsultas] AS
BEGIN
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

CREATE PROCEDURE [dbo].[usp_LeeFormatos]
	(@num_consulta integer) AS
BEGIN
	select *
	from  formatos
	where num_consulta = @num_consulta
	order by nom_columna
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

CREATE PROCEDURE [dbo].[usp_LeePerfiles] AS
BEGIN
	select *
	from perfiles
	order by num_perfil
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

CREATE PROCEDURE [dbo].[usp_LeeProveedores] AS
BEGIN
	select *
	from   Proveedores
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

CREATE PROCEDURE [dbo].[usp_LeeSysConfig] AS
BEGIN
	select *
	from sysconfig
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

CREATE PROCEDURE [dbo].[usp_LeeTipoDatos] AS
BEGIN
	select *
	from  tipo_datos
	order by gls_proveedor, num_tipo_columna_in
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

CREATE PROCEDURE [dbo].[usp_LeeTipoUsuario]
	(@cod_tipo_usuario	nvarchar(12)) AS
BEGIN
	select *
	from   tipo_usuario
	where  cod_tipo_usuario = @cod_tipo_usuario
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

CREATE PROCEDURE [dbo].[usp_LeeTiposUsuarios] AS
BEGIN
	select *
	from   tipo_usuario
	order by cod_tipo_usuario
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

CREATE PROCEDURE [dbo].[usp_LeeUsuarios] AS
BEGIN
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

CREATE PROCEDURE [dbo].[usp_PerfilesPorConsulta]
	(@num_consulta integer) AS
BEGIN
	select distinct p.num_perfil, p.nom_perfil, cp.fec_creacion
	from  cons_perfil cp
	     ,perfiles    p
	where cp.num_consulta = @num_consulta
	and   p.num_perfil    = cp.num_perfil
	order by p.nom_perfil
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

CREATE PROCEDURE [dbo].[usp_PerfilesPorUsuario]
	(@nom_usuario nvarchar(32)) AS
BEGIN
	select p.num_perfil, p.nom_perfil, pu.fec_creacion
	from perf_usuario pu
	    ,perfiles     p
	where pu.nom_usuario = @nom_usuario
	and   p.num_perfil   = pu.num_perfil
	order by p.nom_perfil
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

CREATE PROCEDURE [dbo].[usp_PerfilesSinConsulta]
	(@num_consulta integer) AS
BEGIN
	select distinct *
	from perfiles
	where num_perfil not in (
	select distinct num_perfil
	from cons_perfil
	where num_consulta = @num_consulta
	)
	order by nom_perfil
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

CREATE PROCEDURE [dbo].[usp_PerfilesSinUsuario]
	(@nom_usuario nvarchar(32)) AS
BEGIN
	select *
	from perfiles
	where num_perfil not in (
	select distinct num_perfil
	from perf_usuario
	where nom_usuario = @nom_usuario
	)
	order by nom_perfil
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

CREATE PROCEDURE [dbo].[usp_UsuariosPorConsulta]
	(@num_consulta integer) AS
BEGIN
	select distinct *
	from cons_usuario
	where num_consulta = @num_consulta
	order by nom_usuario
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

CREATE PROCEDURE [dbo].[usp_UsuariosPorPerfil]
	(@num_perfil integer) AS
BEGIN
	select pu.nom_usuario, pu.fec_creacion
	from perfiles     p
	    ,perf_usuario pu
	where p.num_perfil   = @num_perfil
	and   pu.num_perfil  = p.num_perfil
	order by pu.nom_usuario
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

CREATE PROCEDURE [dbo].[usp_UsuariosSinConsulta]
	(@num_consulta integer) AS
BEGIN
	select distinct *
	from usuarios
	where nom_usuario not in (
	select distinct nom_usuario
	from cons_usuario
	where num_consulta = @num_consulta
	)
	order by nom_usuario
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

CREATE PROCEDURE [dbo].[usp_UsuariosSinPerfil]
	(@num_perfil integer) AS
BEGIN
	select *
	from usuarios
	where nom_usuario not in (
	select distinct nom_usuario
	from perf_usuario
	where num_perfil   = @num_perfil
	)
	order by nom_usuario
END

GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

