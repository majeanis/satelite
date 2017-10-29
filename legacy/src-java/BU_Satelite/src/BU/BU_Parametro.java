/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package BU;

import BE.BE_BDatos;
import BE.BE_Parametro;
import DB.DB_BDatos;
import DB.DB_Parametro;
import Ex.BU_Exception;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

/**
 *
 * @author mzavaleta
 */
public class BU_Parametro extends BU_Base {

    public BU_Parametro(BE_BDatos con) throws BU_Exception {
        super(con);
    }

    public BE_Parametro getListaValores(int idQuery, String nomParam) throws BU_Exception {
        Connection conBd = getCn(con);
        BE_BDatos objBDt = null;
        DB_Parametro objDBPar = new DB_Parametro(conBd);
        String query = "";
        try {
            BE_Parametro param = objDBPar.get(idQuery, nomParam);
            if (param == null) {
                return null;
            }
            DB_BDatos objBDBDt = new DB_BDatos(conBd);
            objBDt = objBDBDt.get(param.getIdBDato());
            Class.forName(objBDt.getDriver().getDriver());
            Connection cn = DriverManager.getConnection(objBDt.getUrl(),
                    objBDt.getUsuario(), objBDt.getClave());
            query = param.getQuery();
            PreparedStatement pst = cn.prepareStatement(query,
                    ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY);
            param.setResultSet(pst.executeQuery());
            return param;
        } catch (SQLException ex) {
            throw new BU_Exception("Error al obtener consulta: "
                    + query + " -> " + ex.getMessage());
        } catch (ClassNotFoundException ex) {
            throw new BU_Exception("Driver de Base de Datos no encontrado: "
                    + objBDt.getDriver().getDriver());
        }
    }
}
