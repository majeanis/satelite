package cl.majeanis.satelite.util;

import java.util.Optional;

import org.apache.commons.lang3.StringUtils;

import cl.majeanis.satelite.to.modelo.SesionTO;
import cl.majeanis.satelite.to.modelo.TipoUsuarioTO;
import cl.majeanis.satelite.to.modelo.UsuarioTO;

public final class CheckUtils
{
    static public Resultado checkSesion(SesionTO sesion)
    {
        Resultado r = new ResultadoProceso();
        
        if( sesion == null )
        {
            r.addError("No ha informado la sesión");
            return r;
        }
        
        Resultado ru = checkUsuario(sesion.getUsuario());
        if( !ru.isOk() )
        {
            r.append(ru);
        }
        
        return r;
    }

    static public Resultado checkUsuario(UsuarioTO usuario)
    {
        Resultado r = new ResultadoProceso();

        if( usuario == null )
        {
            r.addError("Usuario no ha sido informado");
            return r;
        } else if( StringUtils.isBlank(usuario.getNombre()))
        {
            r.addError("No se ha informado el nombre del Usuario");
        }

        TipoUsuarioTO tipo = usuario.getTipo();
        if( tipo == null )
        {
            r.addError("Datos del Usuario no tiene el tipo");
            return r;
        } else if( StringUtils.isBlank(tipo.getCodigo()))
        {
            r.addError("No se ha informado el código del tipo de usuario");
        }
        
        return r;
    }
    
    static public Resultado checkSesion(SesionTO sesion, Optional<String> usuario)
    {
        Resultado r = checkSesion(sesion);
        if( usuario == null || !usuario.isPresent())
            return r;
        
        if( r.isOk() )
        {
            if( !usuario.get().equalsIgnoreCase(sesion.getUsuario().getNombre()))
            {
                r.addError("Usuario informado no corresponde al usuario de la sesión");
            }
        }
        
        return r;
    }
}
