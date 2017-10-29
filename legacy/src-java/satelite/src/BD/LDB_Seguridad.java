/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package BD;

import BE.LBE_BDatos;
import BE.LBE_Driver;
import BE.LBE_Opcion;
import BE.LBE_Usuario;
import java.sql.CallableStatement;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Properties;
import oracle.jdbc.OracleTypes;

/**
 *
 * @author mzavaleta
 */
public class LDB_Seguridad extends LDB_Base {

    public LDB_Seguridad(Connection cn, Properties props) {
        this(cn);
        this.props = props;
    }

    public LDB_Seguridad(Connection cn) {
        super(cn);
    }

    public LBE_Usuario login(String usrLogin, String password, int codSistema,
            boolean isChange, String passwordNew) throws SQLException {
        LBE_Usuario usuario = new LBE_Usuario();
        usuario.setUsrLogin(usrLogin);
        CallableStatement cst ;
        String nomSp = "{ call TP_PKG_SEG_NEW.SP_LOGIN(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }";
        cst = cn.prepareCall(nomSp);
        cst.setInt(1, codSistema);
        cst.setString(2, usrLogin);
        cst.setString(3, password);
        cst.registerOutParameter(4, OracleTypes.VARCHAR);
        cst.registerOutParameter(5, OracleTypes.CURSOR);
        //Version
        cst.registerOutParameter(6, OracleTypes.VARCHAR);
        //user_bd
        cst.registerOutParameter(7, OracleTypes.VARCHAR);
        //password bd
        cst.registerOutParameter(8, OracleTypes.VARCHAR);
        //Instancia BD
        cst.registerOutParameter(9, OracleTypes.VARCHAR);
        //Codigo de Tipo de Empleado
        cst.registerOutParameter(10, OracleTypes.VARCHAR);
        //Codigo de Empleado
        cst.registerOutParameter(11, OracleTypes.VARCHAR);
        //Host IP
        cst.setString(12, "");
        //Host name
        cst.setString(13, "");
        //Ip de servidor
        cst.setString(14, "");
        //Flag de cambio de Password
        if (isChange) {
            cst.setString(15, "T");
        } else {
            cst.setString(15, "F");
        }
        //Nuevo password
        cst.setString(16, passwordNew);
        cst.execute();

        ResultSet rs = (ResultSet) cst.getObject(5);
        usuario.setOpcionesAsignadas(new ArrayList<LBE_Opcion>());
        while (rs.next()) {
            LBE_Opcion opc = new LBE_Opcion();
            opc.setCodigo(rs.getInt("cod_opc"));
            opc.setNombre(rs.getString("des_opc"));
            opc.setTipo(rs.getInt("tip_opc"));
            opc.setCodNivel(rs.getString("cod_niv2"));
            opc.setNivel(rs.getInt("nivel"));
            if (rs.getString("asig_TIP").equals("T")) {
                opc.setAsigTipUsr(true);
            } else {
                opc.setAsigTipUsr(false);
            }
            if (rs.getString("asig_usr").equals("T")) {
                opc.setAsigUsr(true);
            } else {
                opc.setAsigUsr(false);
            }
            usuario.getOpcionesAsignadas().add(opc);
        }
        usuario.setNombreCompleto((String) cst.getString(4));
        LBE_BDatos con = new LBE_BDatos();
        LBE_Driver driver =new LBE_Driver();
        driver.setDriver("oracle.jdbc.driver.OracleDriver");
        con.setDriver(driver);
        con.setUrl((String) cst.getString(9));
        con.setUsuario((String) cst.getString(7));
        con.setClave((String) cst.getString(8));
        usuario.setBDatos(con);
        rs.close();
        cst.close();
        return usuario;
    }

    public void cambiaClave(String usrLog, String claveActual, String claveNueva)
            throws SQLException {
        CallableStatement cst;
        String nomSp = "{ call TP_PKG_SEG_NEW.SP_P_CAMBIA_CLAVE(?,?,?) }";
        cst = cn.prepareCall(nomSp);
        cst.setString(1, usrLog);
        cst.setString(2, claveActual);
        cst.setString(3, claveNueva);
        cst.execute();
        cst.close();
    }
}
