package cl.majeanis.satelite.util.tipo;

import java.io.Serializable;

import org.jasypt.util.text.BasicTextEncryptor;

public final class Encrypted implements Serializable, CharSequence, Comparable<Encrypted>
{
    private static final long serialVersionUID = 1L;
    
    private static final BasicTextEncryptor encryptor;
    
    private static final String secret = "1462046bba5d02bbc79953981e829219a928e777d5c1d299bfcbbdd502d3890e";

    private final String value;
    private int          hash;

    static
    {
        synchronized(Encrypted.class)
        {
            encryptor = new BasicTextEncryptor();
            encryptor.setPassword(secret);
        }
    }

    public Encrypted()
    {
        this.value = "";
        this.hash = 0;
    }
    
    public Encrypted(String value)
    {
        if( value == null )
            value = "";

        this.value = encryptor.encrypt(value);
        this.hash = 0;
    }

    @Override
    public int hashCode()
    {
        if( hash!= 0)
            return hash;
        
        hash = String.valueOf(value).hashCode();
        return hash;
    }

    public String plainText()
    {
        return encryptor.decrypt(this.value);
    }

    public String text()
    {
        return this.value;
    }

    @Override
    public int compareTo(Encrypted another)
    {
        if( another == null )
            return 1;
        
        return this.value.compareTo(another.value);
    }

    public int compareTo(String another)
    {
        return compareTo( new Encrypted( another ) );
    }
    
    @Override
    public boolean equals(Object another)
    {
        if( this == another )
            return true;
        
        if( another instanceof Encrypted )
            return this.hashCode() == another.hashCode();
        
        if( another instanceof String )
        {
            Encrypted p = new Encrypted((String) another);
            return this.hashCode() == p.hashCode();
        }
        
        return false;
    }
    
    @Override
    public int length()
    {
        return this.value.length();
    }

    @Override
    public char charAt(int index)
    {
        return this.value.charAt(index);
    }

    @Override
    public CharSequence subSequence(int start, int end)
    {
        return this.value.subSequence(start,  end);
    }
    
    @Override
    public String toString()
    {
        return value;
    }

    static public Encrypted valueOf(String encryptedText)
    {
        String plain = encryptor.decrypt(encryptedText);
        return new Encrypted(plain);
    }
}
