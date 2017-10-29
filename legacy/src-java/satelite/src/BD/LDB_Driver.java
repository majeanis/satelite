/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package BD;


import BE.LBE_Driver;
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
public class LDB_Driver extends LDB_Base {

    public LDB_Driver(Connection cn, Properties props) {
        this(cn);
        this.props=props;
    }

    public LDB_Driver(Connection cn) {
        super(cn);
    }
    public ArrayList<LBE_Driver> getAll() throws SQLException
    {
        ArrayList<LBE_Driver> elems = new ArrayList<LBE_Driver>();
        String query = "select id_drv, nom_drv, drv_cad, url_drv from st_drv s order by s.nom_drv";
        //String query = getQuery("grupos");
        PreparedStatement pst = this.cn.prepareStatement(query);
        ResultSet rs = pst.executeQuery();
        while(rs.next())
        {
            LBE_Driver elem = new LBE_Driver();
            elem.setCodigo(rs.getInt("id_drv"));
            elem.setNombre(rs.getString("nom_drv"));
            elem.setDriver(rs.getString("drv_cad"));
            elem.setUrl(rs.getString("url_drv"));
            elems.add(elem);
        }
        pst.close();
        return elems;
    }
    public LBE_Driver get(int codigo) throws SQLException
    {
        LBE_Driver elem = null;
        String query = "select id_drv, nom_drv, drv_cad from st_drv s where s.id_drv=?";
        //String query = getQuery("grupos");
        PreparedStatement pst = this.cn.prepareStatement(query);
        pst.setInt(1, codigo);
        ResultSet rs = pst.executeQuery();
        if(rs.next())
        {
            elem = new LBE_Driver();
            elem.setCodigo(rs.getInt("id_drv"));
            elem.setNombre(rs.getString("nom_drv"));
            elem.setDriver(rs.getString("drv_cad"));
        }
        pst.close();
        
        return elem;
    }

}
