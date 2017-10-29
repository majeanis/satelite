/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package BU;


import BD.LDB_Seguridad;
import BE.LBE_Usuario;
import BU.Ex.LBU_ClaveCaducoException;
import BU.Ex.LBU_Exception;
import java.sql.SQLException;

/**
 *
 * @author mzavaleta
 */
public class LBU_Seguridad extends LBU_Base {
    public final static int cOPC_MTTO_BASE_DE_DATOS=2;
    public final static int cOPC_MTTO_CONSULTA=3;
    public LBU_Seguridad() throws LBU_Exception {
    }

    public LBE_Usuario login(String usrLogin, String password) throws LBU_Exception, LBU_ClaveCaducoException {
        LBE_Usuario usuario;
        try {
            LDB_Seguridad objDB = new LDB_Seguridad(getCn(), props);
            int codSistema = Integer.parseInt(getProp("codSistema"));
            usuario = objDB.login(usrLogin, password, codSistema, false, "");
            return usuario;
        } catch (SQLException ex) {
            if (ex.getErrorCode() == 20001) {
                throw new LBU_Exception(ex.getMessage());
            } else if (ex.getErrorCode() == 20002) {
                throw new LBU_ClaveCaducoException(ex.getMessage());
            } else {
                throw new LBU_Exception(ex.getMessage());
            }
        }
    }

    public LBE_Usuario cambiaPasswordCaducado(String usrLogin,
            String passwordActual,
            String passwordNuevo) throws LBU_Exception, LBU_ClaveCaducoException {
        LBE_Usuario usuario;
        try {
            LDB_Seguridad objDB = new LDB_Seguridad(getCn(), props);
            int codSistema = Integer.parseInt(getProp("codSistema"));
            usuario = objDB.login(usrLogin, passwordActual, codSistema, true, passwordNuevo);
            return usuario;
        } catch (SQLException ex) {
            if (ex.getErrorCode() == 20001) {
                throw new LBU_Exception(ex.getMessage());
            } else {
                throw new LBU_Exception(ex.getMessage());
            }
        }
    }

    public void cambiaPassword(String usrLogin,
            String passwordActual,
            String passwordNuevo) throws LBU_Exception {
        try {
            LDB_Seguridad objDB = new LDB_Seguridad(getCn(), props);
            objDB.cambiaClave(usrLogin, passwordActual, passwordNuevo);
        } catch (SQLException ex) {
            throw new LBU_Exception(ex.getMessage());
        }
    }
}
