/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package DB;

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
public class DB_Driver extends DB_Base {

    public DB_Driver(Connection cn, Properties props) {
        this(cn);
        this.props=props;
    }

    public DB_Driver(Connection cn) {
        super(cn);
    }
    public ArrayList<BE_Driver> getAll() throws SQLException
    {
        ArrayList<BE_Driver> elems = new ArrayList<BE_Driver>();
        String query = "select id_drv, nom_drv, drv_cad, url_drv from st_drv s order by s.nom_drv";
        //String query = getQuery("grupos");
        PreparedStatement pst = this.cn.prepareStatement(query);
        ResultSet rs = pst.executeQuery();
        while(rs.next())
        {
            BE_Driver elem = new BE_Driver();
            elem.setCodigo(rs.getInt("id_drv"));
            elem.setNombre(rs.getString("nom_drv"));
            elem.setDriver(rs.getString("drv_cad"));
            elem.setUrl(rs.getString("url_drv"));
            elems.add(elem);
        }
        pst.close();
        return elems;
    }
    public BE_Driver get(int codigo) throws SQLException
    {
        BE_Driver elem = null;
        String query = "select id_drv, nom_drv, drv_cad from st_drv s where s.id_drv=?";
        //String query = getQuery("grupos");
        PreparedStatement pst = this.cn.prepareStatement(query);
        pst.setInt(1, codigo);
        ResultSet rs = pst.executeQuery();
        if(rs.next())
        {
            elem = new BE_Driver();
            elem.setCodigo(rs.getInt("id_drv"));
            elem.setNombre(rs.getString("nom_drv"));
            elem.setDriver(rs.getString("drv_cad"));
        }
        pst.close();
        
        return elem;
    }

}
