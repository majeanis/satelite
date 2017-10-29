/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package BU;


import BD.LDB_BDatos;
import BD.LDB_Parametro;
import BE.LBE_BDatos;
import BE.LBE_Parametro;
import BU.Ex.LBU_Exception;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;

/**
 *
 * @author mzavaleta
 */
public class LBU_Parametro extends LBU_Base {

    public LBU_Parametro(LBE_BDatos con) throws LBU_Exception {
        super(con);
    }

    public LBE_Parametro getListaValores(int idQuery, String nomParam) throws LBU_Exception {
        Connection conBd = getCn(con);
        LBE_BDatos objBDt = null;
        LDB_Parametro objDBPar = new LDB_Parametro(conBd);
        String query="";
        try {
            LBE_Parametro param = objDBPar.get(idQuery, nomParam);
            if (param==null)
                return null;
            LDB_BDatos objBDBDt = new LDB_BDatos(conBd);
            objBDt = objBDBDt.get(param.getIdBDato());
            Class.forName(objBDt.getDriver().getDriver());
            Connection cn = DriverManager.getConnection(objBDt.getUrl(),
                    objBDt.getUsuario(), objBDt.getClave());
            query = param.getQuery();
            PreparedStatement pst = cn.prepareStatement(query);
            param.setResultSet(pst.executeQuery());
            return param;
        } catch (SQLException ex) {
            throw new LBU_Exception("Error al obtener consulta: " +
                    query + " -> "  + ex.getMessage());
        }catch (ClassNotFoundException ex) {
            throw new LBU_Exception("Driver de Base de Datos no encontrado: "
                    + objBDt.getDriver().getDriver());
        }
    }
}
