/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package BU;

import BE.BE_BDatos;
import BE.BE_Consulta;
import DB.DB_Consulta;
import Ex.BU_Exception;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

/**
 *
 * @author mzavaleta
 */
public class BU_Consulta extends BU_Base {

    public BU_Consulta(BE_BDatos con) throws BU_Exception {
        super(con);
    }

    public BE_Consulta get(int id) throws BU_Exception {
        try {
            return new DB_Consulta(getCn(con), props).get(id);
        } catch (SQLException ex) {
            throw new BU_Exception("Error al obtener consulta: " + ex.getMessage());
        }
    }

    public ResultSet ejecutar(BE_Consulta be_Consulta, String query,
            String codUsr, boolean esTest) throws BU_Exception {
        DB_Consulta dbConsulta = new DB_Consulta(getCn(con), props);
        try {
            BE_BDatos be_BDatos = be_Consulta.getBe_BDatos();
            Class.forName(be_BDatos.getDriver().getDriver());
            Connection cn = DriverManager.getConnection(be_BDatos.getUrl(),
                    be_BDatos.getUsuario(), be_BDatos.getClave());
            PreparedStatement pst = cn.prepareStatement(query,
            ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY);

            //PreparedStatement pst = cn.prepareStatement(query,
            //        ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);

            ResultSet rs = pst.executeQuery();
            try {
                if (!esTest) {
                    dbConsulta.grabaLog(be_Consulta.getId(), codUsr, true, "");
                    rs.setFetchSize(be_Consulta.getFetchSize());
                }
            } catch (SQLException ex) {
                System.err.println(ex.getMessage());
            }
            return rs;
        } catch (ClassNotFoundException ex) {
            try {
                if (!esTest) {
                    dbConsulta.grabaLog(be_Consulta.getId(), codUsr, false, ex.getMessage());
                }
            } catch (SQLException ex1) {
            }
            throw new BU_Exception("Driver de Base de Datos no encontrado: "
                    + be_Consulta.getBe_BDatos().getDriver());
        } catch (SQLException ex) {
            try {
                if (!esTest) {
                    dbConsulta.grabaLog(be_Consulta.getId(), codUsr, false, ex.getMessage());
                }
            } catch (SQLException ex1) {
            }
            throw new BU_Exception(
                    "Error al ejecutar consulta: " + "\n"
                    + be_Consulta.getBe_BDatos().getUrl() + ": "
                    + ex.getMessage());
        }
    }

    public void add(BE_Consulta consulta, String userLogin) throws BU_Exception {
        try {
            new DB_Consulta(getCn(con), props).add(consulta, userLogin);
        } catch (SQLException ex) {
            throw new BU_Exception("Error al crear consulta: " + ex.getMessage());
        }
    }

    public void save(BE_Consulta consulta, String userLogin) throws BU_Exception {
        try {
            new DB_Consulta(getCn(con), props).save(consulta, userLogin);
        } catch (SQLException ex) {
            throw new BU_Exception("Error al guardar consulta: " + ex.getMessage());
        }
    }

    public ArrayList<BE_Consulta> getLastByUsr(String userLogin) throws BU_Exception {
        try {
            return new DB_Consulta(getCn(con), props).getLastByUsr(userLogin);
        } catch (SQLException ex) {
            throw new BU_Exception("Error al obtener ultimas consultas: " + ex.getMessage());
        }

    }
}
