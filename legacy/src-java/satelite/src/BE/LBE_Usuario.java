/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package BE;

import java.util.ArrayList;

/**
 *
 * @author mzavaleta
 */
public class LBE_Usuario {
    private String usrLogin;
    private String nombreCompleto;
    private String password;
    private LBE_BDatos BDatos;
    protected ArrayList<LBE_Opcion> opcionesAsignadas;

    public ArrayList<LBE_Opcion> getOpcionesAsignadas() {
        return opcionesAsignadas;
    }

    public void setOpcionesAsignadas(ArrayList<LBE_Opcion> opcionesAsignadas) {
        this.opcionesAsignadas = opcionesAsignadas;
    }

    public LBE_BDatos getBDatos() {
        return BDatos;
    }

    public void setBDatos(LBE_BDatos bDatos) {
        this.BDatos = bDatos;
    }

    public String getPassword() {
        return password;
    }

    public void setPassword(String password) {
        this.password = password;
    }

    public String getNombreCompleto() {
        return nombreCompleto;
    }

    public void setNombreCompleto(String nombreCompleto) {
        this.nombreCompleto = nombreCompleto;
    }

    public String getUsrLogin() {
        return usrLogin;
    }

    public void setUsrLogin(String usrLogin) {
        this.usrLogin = usrLogin;
    }
    public boolean validaAccesoAOpcion(int codOpc) {
        for (LBE_Opcion opc : getOpcionesAsignadas()) {
            if (opc.getCodigo() == codOpc) {
                return true;
            }
        }
        return false;
    }

}
