/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package BU;


import BD.LDB_Consulta;
import BE.LBE_BDatos;
import BE.LBE_Consulta;
import BU.Ex.LBU_Exception;
import java.sql.*;
import java.util.ArrayList;

/**
 *
 * @author mzavaleta
 */
public class LBU_Consulta extends LBU_Base {

    public LBU_Consulta(LBE_BDatos con) throws LBU_Exception {
        super(con);
    }

    public LBE_Consulta get(int id) throws LBU_Exception {
        try {
            return new LDB_Consulta(getCn(con), props).get(id);
        } catch (SQLException ex) {
            throw new LBU_Exception("Error al obtener consulta: " + ex.getMessage());
        }
    }

    public ResultSet ejecutar(LBE_Consulta be_Consulta, String query,
            String codUsr, boolean esTest) throws LBU_Exception {
        LDB_Consulta dbConsulta = new LDB_Consulta(getCn(con), props);
        try {
            LBE_BDatos be_BDatos = be_Consulta.getBe_BDatos();
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
            throw new LBU_Exception("Driver de Base de Datos no encontrado: "
                    + be_Consulta.getBe_BDatos().getDriver());
        } catch (SQLException ex) {
            try {
                if (!esTest) {
                    dbConsulta.grabaLog(be_Consulta.getId(), codUsr, false, ex.getMessage());
                }
            } catch (SQLException ex1) {
            }
            throw new LBU_Exception(
                    "Error al ejecutar consulta: " + "\n"
                    + be_Consulta.getBe_BDatos().getUrl() + ": "
                    + ex.getMessage());
        }
    }

    public void add(LBE_Consulta consulta, String userLogin) throws LBU_Exception {
        try {
            new LDB_Consulta(getCn(con), props).add(consulta, userLogin);
        } catch (SQLException ex) {
            throw new LBU_Exception("Error al crear consulta: " + ex.getMessage());
        }
    }

    public void save(LBE_Consulta consulta, String userLogin) throws LBU_Exception {
        try {
            new LDB_Consulta(getCn(con), props).save(consulta, userLogin);
        } catch (SQLException ex) {
            throw new LBU_Exception("Error al guardar consulta: " + ex.getMessage());
        }
    }

    public ArrayList<LBE_Consulta> getLastByUsr(String userLogin) throws LBU_Exception {
        try {
            return new LDB_Consulta(getCn(con), props).getLastByUsr(userLogin);
        } catch (SQLException ex) {
            throw new LBU_Exception("Error al obtener ultimas consultas: " + ex.getMessage());
        }

    }
}
