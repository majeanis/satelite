package cl.majeanis.satelite.bo;

import java.math.BigInteger;
import java.util.List;
import java.util.Optional;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import cl.majeanis.satelite.po.ConsultaPO;
import cl.majeanis.satelite.to.modelo.ConsultaTO;
import cl.majeanis.satelite.to.modelo.SesionTO;
import cl.majeanis.satelite.util.CheckUtils;
import cl.majeanis.satelite.util.Respuesta;
import cl.majeanis.satelite.util.Resultado;
import cl.majeanis.satelite.util.ResultadoProceso;
import cl.majeanis.satelite.util.Utils;

@Service
public class ConsultaBO
{
    private static final Logger logger = LogManager.getLogger(ConsultaBO.class);
    
    @Autowired
    private ConsultaPO consPO;
    
    public Respuesta<List<ConsultaTO>> getList(SesionTO sesion, Optional<String> usuario)
    {
        logger.debug("getList[INI] sesion={}", sesion );

        Resultado rtdo = new ResultadoProceso();

        rtdo = CheckUtils.checkSesion(sesion, usuario);
        if(!rtdo.isOk())
        {
            logger.debug("getList[FIN] errores de validación - {}", rtdo );
            return new Respuesta<>(rtdo);
        }

        /*
         * Si es usuario ADMIN, entonces se buscarán todas las consultas
         */
        List<ConsultaTO> lista = null;
        if( sesion.isAdmin() )
        {
            lista = consPO.getList(Optional.empty(), Optional.empty());            
        } else
        {
            lista = consPO.getList(usuario, Optional.empty());            
        }
       
        logger.debug("getList[FIN] registros retornados={}", Utils.sizeOf(lista));
        return new Respuesta<>(lista);
    }
    
    public Respuesta<ConsultaTO> get(SesionTO sesion, BigInteger consultaId)
    {
        logger.debug("get[INI] consultaId={} sesion={}", consultaId, sesion );
 
        Resultado rtdo = new ResultadoProceso();
        
        rtdo = CheckUtils.checkSesion(sesion);
        if(!rtdo.isOk())
        {
            logger.debug("ge[FIN] errores de validación - {}", rtdo );
            return new Respuesta<>(rtdo);
        }
        
        if( consultaId == null )
        {
            rtdo.addError("Debe informar el número de la consulta");
            return new Respuesta<>(rtdo);
        }

        /*
         * Si es usuario ADMIN, entonces solo se busca por el Id de la Consulta
         */
        ConsultaTO consulta = null;
        if( sesion.isAdmin() )
        {
            consulta = consPO.get(consultaId, Optional.empty());
        } else
        {
            consulta = consPO.get(consultaId, Optional.ofNullable(sesion.getNombreUsuario()));
        }

        logger.debug("get[FIN] registro retornado={}", consulta );
        return new Respuesta<>(consulta);
    }
    
    public Respuesta<ConsultaTO> ejecutar(SesionTO sesion, BigInteger consultaId)
    {
        logger.debug("ejecutar[INI] consultaId={} sesion={}", consultaId, sesion);
        
        return null;
    }
}
