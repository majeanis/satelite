/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package satelite01;

import BE.LBE_Usuario;
import BU.Ex.LBU_Exception;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Properties;

/**
 *
 * @author mzavaleta
 */
public class Constantes {
    public final static String cNOM_APP_PROP="app.properties";
    public final static String cPROP_LOOK_AND_FEEL="look_and_feel";
    public final static String cPROP_WITH_PAN_LEFT="width_pan_left";
    public final static String cPROP_WITH_PAN_LEFT_BD="width_pan_left_bd";
    public final static String cPROP_USR_LOG="usr_login";
    public final static String cPROP_COD_SISTEMA="cod_sistema";
    public static Properties propiedades;
    public static LBE_Usuario UsuarioLogin;

    public static void saveProps() throws LBU_Exception
    {
        try
        {
            try
            {
                propiedades.store(new FileOutputStream(cNOM_APP_PROP),"");
            }catch(FileNotFoundException ex)
            {
                File file = new File(cNOM_APP_PROP);
                file.createNewFile();
                propiedades.store(new FileOutputStream(cNOM_APP_PROP),"");
            }
        }catch(Exception ex)
        {
            throw new LBU_Exception("Error al guardar propiedades :"  +
                    "\n" + ex.getMessage());
        }
    }
}
