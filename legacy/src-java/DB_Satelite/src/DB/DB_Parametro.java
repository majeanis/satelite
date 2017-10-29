/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package DB;

import BE.BE_Parametro;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Properties;

/**
 *
 * @author mzavaleta
 */
public class DB_Parametro extends  DB_Base {

    public DB_Parametro(Connection cn, Properties props) {
        super(cn, props);
    }

    public DB_Parametro(Connection cn) {
        super(cn);
    }

    public BE_Parametro get(int idQuery, String nombre) throws SQLException
    {
        BE_Parametro elem = null;
        String query =
            "select c.id_bdt, t.qry_lis, t.cmp_val\n" +
            "  from ST_CNS_PAR t, st_cns c\n" +
            " where t.id_cns = ?\n" +
            "   and trim(t.nom_par) = ?\n" +
            "   and t.id_cns = c.id_cns";
        //,
        //    ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY
        PreparedStatement pst = cn.prepareStatement(query,
            ResultSet.TYPE_FORWARD_ONLY, ResultSet.CONCUR_READ_ONLY);
        pst.setInt(1, idQuery);
        pst.setString(2, nombre.toUpperCase());
        ResultSet rs = pst.executeQuery();
        rs.setFetchSize(10000);
        if (rs.next()) {
            elem = new BE_Parametro();
            elem.setIdConsulta(idQuery);
            elem.setNombre(nombre);
            elem.setQuery(rs.getString("qry_lis"));
            elem.setIdBDato(rs.getInt("id_bdt"));
            elem.setIdCampoValor(rs.getInt("cmp_val"));
        }
        return elem;
    }

    

}
