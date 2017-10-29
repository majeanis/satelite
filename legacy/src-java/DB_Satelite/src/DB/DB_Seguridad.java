/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package DB;

import BE.BE_BDatos;
import BE.BE_Driver;
import BE.BE_Opcion;
import BE.BE_Usuario;
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
public class DB_Seguridad extends DB_Base {

    public DB_Seguridad(Connection cn, Properties props) {
        this(cn);
        this.props = props;
    }

    public DB_Seguridad(Connection cn) {
        super(cn);
    }

    public BE_Usuario login(String usrLogin, String password, int codSistema,
            boolean isChange, String passwordNew) throws SQLException {
        BE_Usuario usuario = new BE_Usuario();
        usuario.setUsrLogin(usrLogin);
        CallableStatement cst = null;
        String nomSp = null;
        //try {
        nomSp = "{ call TP_PKG_SEG_NEW.SP_LOGIN(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }";
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
        usuario.setOpcionesAsignadas(new ArrayList<BE_Opcion>());
        while (rs.next()) {
            BE_Opcion opc = new BE_Opcion();
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
        BE_BDatos con = new BE_BDatos();
        BE_Driver driver =new BE_Driver();
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
        CallableStatement cst = null;
        String nomSp = null;
        nomSp = "{ call TP_PKG_SEG_NEW.SP_P_CAMBIA_CLAVE(?,?,?) }";
        cst = cn.prepareCall(nomSp);
        cst.setString(1, usrLog);
        cst.setString(2, claveActual);
        cst.setString(3, claveNueva);
        cst.execute();
        cst.close();
    }
}
