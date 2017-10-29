/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package DB;

import BE.BE_BDatos;
import BE.BE_Driver;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Properties;

/**
 *
 * @author mzavaleta
 */
public class DB_BDatos extends DB_Base {

    public DB_BDatos(Connection cn, Properties props) {
        super(cn, props);
    }
    public DB_BDatos(Connection cn) {
        super(cn);
    }
    public ArrayList<BE_BDatos> getAll() throws SQLException {
        ArrayList<BE_BDatos> elems = new ArrayList<BE_BDatos>();
        String query = "select id_bdt, nom_bdt, url_bdt, usr_bdt, nom_drv," +
                " url_bdt,psw_bdt, drv_cad, id_drv " +
                " from st_vw_bdt order by id_bdt";
        PreparedStatement pst = this.cn.prepareStatement(query);
        ResultSet rs = pst.executeQuery();
        while (rs.next()) {
            BE_BDatos elem = new BE_BDatos();
            elem.setCodigo(rs.getInt("id_bdt"));
            elem.setNombre(rs.getString("nom_bdt"));
            elem.setUrl(rs.getString("url_bdt"));
            elem.setUsuario(rs.getString("usr_bdt"));
            elem.setClave(rs.getString("psw_bdt"));

            BE_Driver Driver = new BE_Driver();
            Driver.setCodigo(rs.getInt("id_drv"));
            Driver.setNombre(rs.getString("nom_drv"));
            Driver.setDriver(rs.getString("drv_cad"));
            elem.setDriver(Driver);

            elems.add(elem);
        }
        pst.close();
        return elems;
    }

    public BE_BDatos get(int codigo) throws SQLException {
        BE_BDatos elem = null;
        String query = "select id_bdt, nom_bdt, url_bdt, usr_bdt, nom_drv," +
                "b.id_drv,psw_bdt, drv_cad from st_vw_bdt b where b.id_bdt=?";
        PreparedStatement pst = this.cn.prepareStatement(query);
        pst.setInt(1, codigo);
        ResultSet rs = pst.executeQuery();
        if (rs.next()) {
            elem = new BE_BDatos();
            elem.setCodigo(rs.getInt("id_bdt"));
            elem.setNombre(rs.getString("nom_bdt"));
            elem.setUrl(rs.getString("url_bdt"));
            elem.setUsuario(rs.getString("usr_bdt"));
            elem.setClave(rs.getString("psw_bdt"));

            BE_Driver Driver = new BE_Driver();
            Driver.setCodigo(rs.getInt("id_drv"));
            Driver.setNombre(rs.getString("nom_drv"));
            Driver.setDriver(rs.getString("drv_cad"));
            elem.setDriver(Driver);
        }
        pst.close();
        return elem;
    }

    public void add(BE_BDatos bDatos, String usrLogin) throws SQLException {
        String comando =
                "insert into st_bdt\n"
                + "  (id_drv, id_bdt, nom_bdt, url_bdt, usr_bdt, psw_bdt, fec_cre, usr_cre)\n"
                + "values\n"
                + "  (?, ST_SQ_BDT.NEXTVAL, ?, ?, ?, ?, SYSDATE, ?)";
        PreparedStatement pst = this.cn.prepareStatement(comando);
        pst.setInt(1, bDatos.getDriver().getCodigo());
        pst.setString(2, bDatos.getNombre());
        pst.setString(3, bDatos.getUrl());
        pst.setString(4, bDatos.getUsuario());
        pst.setString(5, bDatos.getClave());
        pst.setString(6, usrLogin);
        pst.executeUpdate();
        pst.close();
    }

    public void save(BE_BDatos bDatos, String usrLogin) throws SQLException {
        String comando =
                "update st_bdt b\n"
                + "   set b.id_drv  = ?,\n"
                + "       b.nom_bdt = ?,\n"
                + "       b.url_bdt = ?,\n"
                + "       b.usr_bdt = ?,\n"
                + "       b.psw_bdt = ?,\n"
                + "       b.fec_mod = sysdate,\n"
                + "       b.usr_mod = ?\n"
                + " where b.id_bdt = ?";
        PreparedStatement pst = this.cn.prepareStatement(comando);
        pst.setInt(1, bDatos.getDriver().getCodigo());
        pst.setString(2, bDatos.getNombre());
        pst.setString(3, bDatos.getUrl());
        pst.setString(4, bDatos.getUsuario());
        pst.setString(5, bDatos.getClave());
        pst.setString(6, usrLogin);
        pst.setInt(7, bDatos.getCodigo());
        pst.executeUpdate();
        pst.close();
    }
}
