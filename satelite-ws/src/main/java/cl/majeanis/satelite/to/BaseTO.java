package cl.majeanis.satelite.to;

import org.apache.commons.lang3.builder.CompareToBuilder;
import org.apache.commons.lang3.builder.EqualsBuilder;
import org.apache.commons.lang3.builder.HashCodeBuilder;

import cl.majeanis.satelite.util.ToStringUtils;

public class BaseTO implements ObjetoTO
{
    private static final long serialVersionUID = 1L;
    
    @Override
    public String toString()
    {
        return ToStringUtils.toString(this);
    }

    @Override
    public int compareTo(ObjetoTO obj)
    {
        return CompareToBuilder.reflectionCompare(this,obj);
    }

    @Override
    public boolean equals(Object obj)
    {
        if( obj == null )
            return false;
        
        if( this == obj )
            return true;
        
        if( obj instanceof ObjetoTO )
            return this.hashCode() == obj.hashCode();
        
        return EqualsBuilder.reflectionEquals(this, obj);
    }

    @Override
    public int hashCode()
    {
        return HashCodeBuilder.reflectionHashCode(this);
    }
}
