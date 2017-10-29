/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package BU;

import BE.BE_Usuario;
import DB.DB_Seguridad;
import Ex.BU_ClaveCaducoException;
import Ex.BU_Exception;
import java.sql.SQLException;

/**
 *
 * @author mzavaleta
 */
public class BU_Seguridad extends BU_Base {
    public final static int cOPC_MTTO_BASE_DE_DATOS=2;
    public final static int cOPC_MTTO_CONSULTA=3;
    public BU_Seguridad() throws BU_Exception {
    }

    public BE_Usuario login(String usrLogin, String password) throws BU_Exception, BU_ClaveCaducoException {
        BE_Usuario usuario;
        try {
            DB_Seguridad objDB = new DB_Seguridad(getCn(), props);
            int codSistema = Integer.parseInt(getProp("codSistema"));
            usuario = objDB.login(usrLogin, password, codSistema, false, "");
            return usuario;
        } catch (SQLException ex) {
            if (ex.getErrorCode() == 20001) {
                throw new BU_Exception(ex.getMessage());
            } else if (ex.getErrorCode() == 20002) {
                throw new BU_ClaveCaducoException(ex.getMessage());
            } else {
                throw new BU_Exception(ex.getMessage());
            }
        }
    }

    public BE_Usuario cambiaPasswordCaducado(String usrLogin,
            String passwordActual,
            String passwordNuevo) throws BU_Exception, BU_ClaveCaducoException {
        BE_Usuario usuario;
        try {
            DB_Seguridad objDB = new DB_Seguridad(getCn(), props);
            int codSistema = Integer.parseInt(getProp("codSistema"));
            usuario = objDB.login(usrLogin, passwordActual, codSistema, true, passwordNuevo);
            return usuario;
        } catch (SQLException ex) {
            if (ex.getErrorCode() == 20001) {
                throw new BU_Exception(ex.getMessage());
            } else {
                throw new BU_Exception(ex.getMessage());
            }
        }
    }

    public void cambiaPassword(String usrLogin,
            String passwordActual,
            String passwordNuevo) throws BU_Exception {
        try {
            DB_Seguridad objDB = new DB_Seguridad(getCn(), props);
            objDB.cambiaClave(usrLogin, passwordActual, passwordNuevo);
        } catch (SQLException ex) {
            throw new BU_Exception(ex.getMessage());
        }
    }
}
