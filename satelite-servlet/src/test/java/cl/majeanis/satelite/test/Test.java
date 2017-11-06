package cl.majeanis.satelite.test;

import cl.majeanis.satelite.util.Resultado;
import cl.majeanis.satelite.util.ResultadoProceso;

public class Test
{

    public static void main(String[] args)
    {
        Resultado r = new ResultadoProceso();
        
        r.addMensaje("un mensaje");
        r.addMensaje("dos mensaje");
        r.addError( new Exception("hola") );
        
        System.out.println( r );
    }

}
