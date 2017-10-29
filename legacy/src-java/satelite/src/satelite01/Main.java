/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package satelite01;

import java.awt.Frame;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import logger.Log;

/**
 *
 * @author mzavaleta
 */
public class Main {

    //static public Logger logger = Logger.getLogger(Main.class);

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException, InstantiationException, IllegalAccessException {
        // TODO code application logic here
        //PropertyConfigurator.configure("config.properties");
        MdiMain mdiMain = new MdiMain();
        Properties prop = new Properties();
        try {
            try {
                InputStream is = new FileInputStream(Constantes.cNOM_APP_PROP);
                prop.load(is);
                //PropertyConfigurator.configure(prop);
                is.close();
                //logger.debug("Inicio de aplicacion");

            } catch (FileNotFoundException ex) {
            }
            Constantes.propiedades = prop;
            Log.debug("Inicio");
            String look_and_feel = prop.getProperty(Constantes.cPROP_LOOK_AND_FEEL);
            if (look_and_feel != null) {
                UIManager.setLookAndFeel(look_and_feel);
                SwingUtilities.updateComponentTreeUI(mdiMain);
            }
            String width_pan_left = prop.getProperty(Constantes.cPROP_WITH_PAN_LEFT);
            if (width_pan_left != null) {
                mdiMain.setDividerLocation(Integer.parseInt(width_pan_left));
            }
            mdiMain.setExtendedState(Frame.MAXIMIZED_BOTH);
            mdiMain.setVisible(true);
            jfLogin login = new jfLogin(mdiMain, true);
            login.setLocationRelativeTo(mdiMain);
            login.setVisible(true);
            Log.debug("inicio carga de arbol");
            mdiMain.cargaArbol();
            Log.debug("Fin carga de arbol");
        } catch (ClassNotFoundException ex) {
        } catch (UnsupportedLookAndFeelException ex) {
        }
    }
}
