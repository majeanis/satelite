/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package BE;

import java.sql.ResultSet;

/**
 *
 * @author mzavaleta
 */
public class LBE_Parametro {
    private int idConsulta;
    private int idBDato;
    private int idCampoValor;
    private ResultSet resultSet;

    public ResultSet getResultSet() {
        return resultSet;
    }

    public void setResultSet(ResultSet resultSet) {
        this.resultSet = resultSet;
    }


    public int getIdCampoValor() {
        return idCampoValor;
    }

    public void setIdCampoValor(int idCampoValor) {
        this.idCampoValor = idCampoValor;
    }


    public int getIdBDato() {
        return idBDato;
    }

    public void setIdBDato(int idBDato) {
        this.idBDato = idBDato;
    }

    public String getValor() {
        return Valor;
    }

    public void setValor(String Valor) {
        this.Valor = Valor;
    }

    public int getIdConsulta() {
        return idConsulta;
    }

    public void setIdConsulta(int idConsulta) {
        this.idConsulta = idConsulta;
    }

    public String getNombre() {
        return nombre;
    }

    public void setNombre(String nombre) {
        this.nombre = nombre;
    }

    public String getQuery() {
        return query;
    }

    public void setQuery(String query) {
        this.query = query;
    }
    private String query;
    private String nombre;
    private String Valor;
}
