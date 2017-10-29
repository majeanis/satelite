/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package DB;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Properties;

/**
 *
 * @author mzavaleta
 */
public class DB_Base {

    protected Connection cn;
    protected Properties props;
    protected String getQuery(String label)
    {
        String nombre = props.getProperty("db");
        return props.getProperty(nombre + ".query." + label,"");
    }

    public DB_Base(Connection cn) {
        this.cn=cn;
    }
    public DB_Base(Connection cn, Properties props) {
        this(cn);
        this.props= props;
    }
    protected int getSeq(String query) throws SQLException
    {
        int resultado=0;
        PreparedStatement pst = cn.prepareStatement(query);
        ResultSet rs = pst.executeQuery();
        if(rs.next())
        {
            resultado = rs.getInt(1);
        }
        return resultado;
    }
}
