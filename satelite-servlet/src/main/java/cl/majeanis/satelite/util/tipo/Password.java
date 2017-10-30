package cl.majeanis.satelite.util.tipo;

import java.io.Serializable;

public final class Password implements Serializable, CharSequence, Comparable<Password>
{
    private static final long serialVersionUID = 1L;
    
    private final char        value[];
    private final int         count;
    private int               hash;

    public Password()
    {
        this.value = new char[0];
        this.count = 0;
        this.hash = 0;
    }
    
    public Password(String value)
    {
        this.value = encode(value).toCharArray();
        this.count = this.value.length;
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

    @Override
    public int compareTo(Password another)
    {
        int len1 = this.count;
        int len2 = another.count;
        int maxIndex = Math.min(len1, len2);

        for (int i = 0; i < maxIndex; i++)
        {
            char c1 = this.value[i];
            char c2 = another.value[i];

            if (c1 != c2)
            {
                return c1 - c2;
            }
        }

        return (len1 - len2);
    }

    public int compareTo(String another)
    {
        return compareTo( new Password( another ) );
    }
    
    @Override
    public boolean equals(Object another)
    {
        if( this == another )
            return true;
        
        if( another instanceof Password )
            return this.hashCode() == another.hashCode();
        
        if( another instanceof String )
        {
            String s = encode((String) another);
            return this.hashCode() == s.hashCode();
        }
        
        return false;
    }
    
    @Override
    public int length()
    {
        return this.count;
    }

    @Override
    public char charAt(int index)
    {
        return this.value[index];
    }

    @Override
    public CharSequence subSequence(int start, int end)
    {
        // TODO Auto-generated method stub
        return null;
    }
    
    private String encode(String plainText)
    {
        if( plainText == null) return "";
        return plainText;
    }
    
    public String plainText()
    {
        return String.valueOf(value);
    }
}
