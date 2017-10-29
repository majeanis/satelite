/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package BE;


import BE.Ex.LBE_Exception;
import java.util.ArrayList;
import java.util.Date;

/**
 *
 * @author mzavaleta
 */
public class LBE_Consulta {

    protected int id;
    private Date lastAcceso;
    private String titulo;
    protected LBE_BDatos be_BDatos;
    private LBE_Grupo Grupo;
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
    

    public LBE_BDatos getBe_BDatos() {
        return be_BDatos;
    }

    public void setBe_BDatos(LBE_BDatos be_BDatos) {
        this.be_BDatos = be_BDatos;
    }


    public LBE_Grupo getGrupo() {
        return Grupo;
    }

    public void setGrupo(LBE_Grupo Grupo) {
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

    public ArrayList<String> parametros() throws LBE_Exception {
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
                    throw new LBE_Exception("Query incorrecto");
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
