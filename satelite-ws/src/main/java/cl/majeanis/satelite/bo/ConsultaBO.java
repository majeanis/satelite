package cl.majeanis.satelite.bo;

import java.util.List;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import cl.majeanis.satelite.po.ConsultaPO;
import cl.majeanis.satelite.to.modelo.ConsultaTO;
import cl.majeanis.satelite.to.modelo.SesionTO;
import cl.majeanis.satelite.to.modelo.TipoUsuarioTO;
import cl.majeanis.satelite.to.modelo.UsuarioTO;
import cl.majeanis.satelite.util.Respuesta;
import cl.majeanis.satelite.util.Resultado;
import cl.majeanis.satelite.util.ResultadoProceso;
import cl.majeanis.satelite.util.Utils;
import cl.majeanis.satelite.util.tipo.TipoUsuario;

@Service
public class ConsultaBO
{
    private static final Logger logger = LogManager.getLogger(ConsultaBO.class);
    
    @Autowired
    private ConsultaPO consPO;
    
    public Respuesta<List<ConsultaTO>> getList(SesionTO sesion)
    {
        logger.debug("getList[INI] sesion={}", sesion );
        
        Resultado rtdo = new ResultadoProceso();
        if( sesion == null )
        {
            rtdo.addError("Debe informar datos de la sesión");
            logger.debug("getList[FIN] no ha informado la sesión" );
            return new Respuesta<>(rtdo);
        }
        
        UsuarioTO usuario = sesion.getUsuario();
        if( usuario == null )
        {
            rtdo.addError("Sesión no tiene información del usaurio");
            logger.debug("getList[FIN] sesión no tiene datos del usuario={}", sesion );
            return new Respuesta<>(rtdo);
        }
        
        TipoUsuarioTO tipo = usuario.getTipo();
        if( tipo == null )
        {
            rtdo.addError("Datos del Usuario no tiene el tipo");
            logger.debug("getList[FIN] usuario tiene definido el tipo={}", sesion );
            return new Respuesta<>(rtdo);
        }
        
        if( !TipoUsuario.ADMIN.name().equalsIgnoreCase(tipo.getCodigo()))
        {
            rtdo.addError("Solo usuario Administrador puede listar todas las consultas");
            logger.debug("getList[FIN] usuario no es ADMIN={}", sesion );
            return new Respuesta<>(rtdo);
        }
        
        List<ConsultaTO> lista = consPO.getList(null);
        logger.debug("getList[FIN] registros retornados={}", Utils.sizeOf(lista));
        return new Respuesta<>(lista);
    }
    
//    public Respuesta<List<ConsultaTO>> getList(SesionTO sesion, String usuario)
//    {
//        logger.debug("getList[INI] usuario={} sesion={}", usuario, sesion );
//        
//        Resultado rtdo = new ResultadoProceso();
//        if( StringUtils.isBlank(usuario) )
//        {
//            rtdo.addError("Debe informar nombre del usuario");
//        } else if( sesion != null && 
//                   sesion.getUsuario() != null && 
//                  !usuario.equals(sesion.getUsuario().getNombre()))
//        {
//            rtdo.addError("Nombre de usuario no corresponde al de la sesión" );
//        }
//        
//        if(!rtdo.isOk())
//        {
//            logger.debug("getList[FIN] errores de validación={}", rtdo);
//            return new Respuesta<>(rtdo);
//        }
//        
//        List<ConsultaTO> lista = consPO.getList( null, usuario );
//        
//        logger.debug("getList[FIN] registros retornados={}", Utils.sizeOf(lista) );
//        return new Respuesta<>(lista);
//    }
}
