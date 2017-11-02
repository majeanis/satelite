package cl.majeanis.satelite.util;

import java.io.Serializable;
import java.util.Optional;

import cl.majeanis.satelite.to.ObjetoTO;

public class Respuesta<T extends ObjetoTO> implements Serializable
{
    private static final long serialVersionUID = 1L;

    private final Resultado result;
    private final Optional<T> content;
    
    public Respuesta(Resultado result, T content)
    {
        this.result = (result == null ? new ResultadoProceso(): result); 
        this.content = Optional.ofNullable(content);
    }

    public Respuesta(Resultado result)
    {
        this(result, null);
    }

    public Respuesta(T content)
    {
        this(null, content);
    }

    public Respuesta(Exception exception)
    {
        Resultado r = new ResultadoProceso();
        r.addError(exception);
        this.result = r;
        this.content = Optional.empty();
    }

    public Resultado getResult()
    {
        return result;
    }

    public Optional<T> getContent()
    {
        return content;
    }
    
    public boolean isOk()
    {
        return result.isOk();
    }
    
    public boolean isContentOk()
    {
        return content.isPresent() && result.isOk();
    }

    public String toString()
    {
        return "Respuesta[result=" + ToStringUtils.toString(this) + ",content=" + content.isPresent() + "]";
    }
}
