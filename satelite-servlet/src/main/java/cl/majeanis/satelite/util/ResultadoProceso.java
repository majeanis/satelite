package cl.majeanis.satelite.util;

import java.time.Instant;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class ResultadoProceso implements Resultado
{
    private static final long serialVersionUID = 1L;

    private final List<String>    mensajes;
    private final List<String>    errores;
    private final List<Exception> exceptions;

    public ResultadoProceso()
    {
        this.mensajes = new ArrayList<>();
        this.errores = new ArrayList<>();
        this.exceptions = new ArrayList<>();
    }

    public ResultadoProceso(Resultado another)
    {
        this();
        this.append(another);
    }

    @Override
    public boolean hasErrors()
    {
        return !this.errores.isEmpty();
    }

    @Override
    public boolean hasExceptions()
    {
        return !this.exceptions.isEmpty();
    }

    @Override
    public void addMensaje(String format, Object... args)
    {
        this.mensajes.add(String.format(format, args));
    }

    @Override
    public void addError(String format, Object... args)
    {
        this.errores.add(String.format(format, args));
    }

    @Override
    public void addError(Exception exception)
    {
        addError("Error no esperado [id=%1$d,name=%2$s,message=%3$s]", 
                Instant.now().hashCode(),
                exception.getClass().getSimpleName(), 
                exception.getMessage());
        this.exceptions.add(exception);
    }
    
    @Override
    public void append(Resultado another)
    {
        this.mensajes.addAll(another.getMensajes());
        this.errores.addAll(another.getErrores());
        this.exceptions.addAll(another.getExceptions());
    }

    @Override
    public List<String> getMensajes()
    {
        return (List<String>) Collections.unmodifiableCollection(this.mensajes);
    }

    @Override
    public List<String> getErrores()
    {
        return (List<String>) Collections.unmodifiableCollection(this.errores);
    }

    @Override
    public List<Exception> getExceptions()
    {
        return (List<Exception>) Collections.unmodifiableCollection(this.exceptions);
    }
    
    @Override
    public String toString()
    {
        return ToStringUtils.toString(this);
    }
}
