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
public class BE_Usuario {
    private String usrLogin;
    private String nombreCompleto;
    private String password;
    private BE_BDatos BDatos;
    protected ArrayList<BE_Opcion> opcionesAsignadas;

    public ArrayList<BE_Opcion> getOpcionesAsignadas() {
        return opcionesAsignadas;
    }

    public void setOpcionesAsignadas(ArrayList<BE_Opcion> opcionesAsignadas) {
        this.opcionesAsignadas = opcionesAsignadas;
    }

    public BE_BDatos getBDatos() {
        return BDatos;
    }

    public void setBDatos(BE_BDatos bDatos) {
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
        for (BE_Opcion opc : getOpcionesAsignadas()) {
            if (opc.getCodigo() == codOpc) {
                return true;
            }
        }
        return false;
    }

}
