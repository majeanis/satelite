/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package BE;

import Ex.BE_Exception;
import java.util.ArrayList;
import java.util.Date;

/**
 *
 * @author mzavaleta
 */
public class BE_Consulta {

    protected int id;
    private Date lastAcceso;
    private String titulo;
    protected BE_BDatos be_BDatos;
    private BE_Grupo Grupo;
    private String query;
    protected int idOpcionSegur;
    int fetchSize=10000;

    public int getFetchSize() {
        return fetchSize;
    }

    public void setFetchSize(int fetchSize) {
        this.fetchSize = fetchSize;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public Date getLastAcceso() {
        return lastAcceso;
    }

    public void setLastAcceso(Date lastAcceso) {
        this.lastAcceso = lastAcceso;
    }
    

    public BE_BDatos getBe_BDatos() {
        return be_BDatos;
    }

    public void setBe_BDatos(BE_BDatos be_BDatos) {
        this.be_BDatos = be_BDatos;
    }


    public BE_Grupo getGrupo() {
        return Grupo;
    }

    public void setGrupo(BE_Grupo Grupo) {
        this.Grupo = Grupo;
    }

    public String getQuery() {
        return query;
    }

    public void setQuery(String query) {
        this.query = query;
    }

    public String getTitulo() {
        return titulo;
    }

    public void setTitulo(String titulo) {
        this.titulo = titulo;
    }

    @Override
    public String toString() {
        return this.titulo;
    }

    public ArrayList<String> parametros() throws BE_Exception {
        ArrayList<String> params = new ArrayList<String>();
        int pos = 0;
        boolean sigue = true;
        while (sigue) {
            int result1 = query.indexOf("@", pos);
            int result2 = 0;
            if (result1 == -1) {
                sigue = false;
            } else {
                result2 = query.indexOf("@", result1 + 1);
                if (result2 == -1) {
                    throw new BE_Exception("Query incorrecto");
                }
                String parametro = query.substring(result1 + 1, result2);
                //Verificamos si existe el parámetro
                boolean existe = false;
                for (String par : params) {
                    if (par.equalsIgnoreCase(parametro))
                        existe=true;
                }
                if (!existe)
                    params.add(parametro);
            }
            pos = result2 + 1;
        }
        return params;
    }

    public int getIdOpcionSegur() {
        return idOpcionSegur;
    }

    public void setIdOpcionSegur(int idOpcionSegur) {
        this.idOpcionSegur = idOpcionSegur;
    }
}
