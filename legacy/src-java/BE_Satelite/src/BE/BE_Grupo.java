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
public class BE_Grupo {
    public BE_Grupo()
    {
        consultas = new ArrayList<BE_Consulta>();
    }
    protected int id;
    protected String nombre;
    protected ArrayList<BE_Consulta> consultas;
    protected int idOpcionSegur;

    public int getIdOpcionSegur() {
        return idOpcionSegur;
    }

    public void setIdOpcionSegur(int idOpcionSegur) {
        this.idOpcionSegur = idOpcionSegur;
    }

    public ArrayList<BE_Consulta> getConsultas() {
        return consultas;
    }

    public void setConsultas(ArrayList<BE_Consulta> consultas) {
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
