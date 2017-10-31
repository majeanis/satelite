package cl.majeanis.satelite.util.tipo;

import cl.majeanis.satelite.to.modelo.DriverTO;

public enum DriverJdbc
{
    ORACLE(1),
    SQL_SERVER(2)
    ;
    
    private int id;
    
    DriverJdbc(int id)
    {
        this.id = id;
    }
    
    static public DriverJdbc of(Integer id)
    {
        for(DriverJdbc v: DriverJdbc.values())
        {
            if( v.id == id )
                return v;
        }
        
        return null;
    }

    static public DriverJdbc of(DriverTO driver)
    {
        if( driver == null ) return null;
        return of(driver.getId());
    }
}
