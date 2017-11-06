package cl.majeanis.satelite.util.tipo;

import java.time.LocalTime;

import cl.majeanis.satelite.util.ToStringUtils;

public class RangoHora
{
    private LocalTime desde;
    private LocalTime hasta;
    
    public RangoHora(LocalTime desde, LocalTime hasta)
    {
        this.desde = desde;
        this.hasta = hasta;
    }
    
    public LocalTime getDesde()
    {
        return this.desde;
    }
    
    public LocalTime getHasta()
    {
        return this.hasta;
    }
    
    public void setDesde(LocalTime desde)
    {
        this.desde = desde;
    }
    
    public void setHasta(LocalTime hasta)
    {
        this.hasta = hasta;
    }
    
    @Override
    public String toString()
    {
        return ToStringUtils.toString(this);
    }
}
