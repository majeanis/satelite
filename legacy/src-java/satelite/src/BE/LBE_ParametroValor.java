/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package BE;

/**
 *
 * @author mzavaleta
 */
public class LBE_ParametroValor {
    private String valor;
    private String descripcion;

    public String getDescripcion() {
        return descripcion;
    }

    public void setDescripcion(String descripcion) {
        this.descripcion = descripcion;
    }

    public String getValor() {
        return valor;
    }

    public void setValor(String valor) {
        this.valor = valor;
    }
    @Override
    public String toString()
    {
        return this.descripcion;
    }
}
