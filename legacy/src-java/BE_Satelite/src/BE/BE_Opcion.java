/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package BE;

/**
 *
 * @author mzavaleta
 */
public class BE_Opcion {
    protected int codigo;

    public boolean isAsigTipUsr() {
        return asigTipUsr;
    }

    public void setAsigTipUsr(boolean asigTipUsr) {
        this.asigTipUsr = asigTipUsr;
    }

    public boolean isAsigUsr() {
        return asigUsr;
    }

    public void setAsigUsr(boolean asigUsr) {
        this.asigUsr = asigUsr;
    }

    public String getCodNivel() {
        return codNivel;
    }

    public void setCodNivel(String codNivel) {
        this.codNivel = codNivel;
    }

    public int getCodigo() {
        return codigo;
    }

    public void setCodigo(int codigo) {
        this.codigo = codigo;
    }

    public int getNivel() {
        return nivel;
    }

    public void setNivel(int nivel) {
        this.nivel = nivel;
    }

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }

    public int getTipo() {
        return tipo;
    }

    public void setTipo(int tipo) {
        this.tipo = tipo;
    }
    protected String nombre;
    protected int tipo;
    protected String codNivel;
    protected int nivel;
    protected boolean asigTipUsr;
    protected boolean asigUsr;



}
