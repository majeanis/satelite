package cl.majeanis.satelite.to;

import java.math.BigInteger;

public abstract class PersistibleTO extends BaseTO
{
    private static final long serialVersionUID = 1L;

    private BigInteger id;

    public BigInteger getId()
    {
        return id;
    }

    public void setId(BigInteger id)
    {
        this.id = id;
    }
    
    public boolean isIdBlank()
    {
        return id == null || id.equals(0);
    }
}
