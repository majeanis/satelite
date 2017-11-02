package cl.majeanis.satelite.util;

import java.io.Serializable;
import java.util.List;

public interface Resultado extends Serializable
{
    default boolean isOk()
    {
        return !hasErrors() && !hasExceptions();
    }

    public boolean hasExceptions();

    public boolean hasErrors();

    public List<String> getMensajes();

    public List<String> getErrores();

    public List<Exception> getExceptions();

    public void addMensaje(String format, Object... args);

    public void addError(String format, Object... args);

    public void addError(Exception exception);

    public void append(Resultado another);
}
