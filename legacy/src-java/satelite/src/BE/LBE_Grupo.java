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
public class LBE_Grupo {
    public LBE_Grupo()
    {
        consultas = new ArrayList<LBE_Consulta>();
    }
    protected int id;
    protected String nombre;
    protected ArrayList<LBE_Consulta> consultas;
    protected int idOpcionSegur;

    public int getIdOpcionSegur() {
        return idOpcionSegur;
    }

    public void setIdOpcionSegur(int idOpcionSegur) {
        this.idOpcionSegur = idOpcionSegur;
    }

    public ArrayList<LBE_Consulta> getConsultas() {
        return consultas;
    }

    public void setConsultas(ArrayList<LBE_Consulta> consultas) {
        this.consultas = consultas;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }
    @Override
    public String toString()
    {
        return this.nombre;
    }
}
