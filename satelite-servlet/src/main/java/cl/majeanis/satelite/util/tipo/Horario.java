package cl.majeanis.satelite.util.tipo;

import cl.majeanis.satelite.util.ToStringUtils;

public class Horario
{
    private final Boolean diaCompleto;
    private final RangoHora periodo1;
    private final RangoHora periodo2;
    
    public Horario(Boolean diaCompleto)
    {
        this.diaCompleto = diaCompleto;
        this.periodo1 = null;
        this.periodo2 = null;
    }
    
    public Horario(RangoHora periodo)
    {
        this.diaCompleto = false;
        this.periodo1 = periodo;
        this.periodo2 = null;
    }

    public Horario(RangoHora periodo1, RangoHora periodo2)
    {
        this.diaCompleto = false;
        this.periodo1 = periodo1;
        this.periodo2 = periodo2;
    }

    public Boolean getDiaCompleto()
    {
        return diaCompleto;
    }

    public RangoHora getPeriodo1()
    {
        return periodo1;
    }

    public RangoHora getPeriodo2()
    {
        return periodo2;
    }
    
    @Override
    public String toString()
    {
        return ToStringUtils.toString(this);
    }
}
