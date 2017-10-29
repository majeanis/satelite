/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package BD;

import BE.LBE_BDatos;
import BE.LBE_Consulta;
import BE.LBE_Driver;
import BE.LBE_Grupo;
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
public class LDB_Consulta extends LDB_Base {

    public LDB_Consulta(Connection cn) {
        super(cn);
    }

    public LDB_Consulta(Connection cn, Properties props) {
        this(cn);
        this.props = props;
    }

    public LBE_Consulta get(int id) throws SQLException {
        LBE_Consulta be_Consulta = null;
        String query = getQuery("getConsulta");
        PreparedStatement pst = cn.prepareStatement(query);
        pst.setInt(1, id);
        ResultSet rs = pst.executeQuery();
        if (rs.next()) {
            be_Consulta = new LBE_Consulta();
            be_Consulta.setId(id);
            be_Consulta.setTitulo(rs.getString(1));
            String qry = rs.getString(2);
            if (rs.getObject(8) != null) {
                qry = qry + rs.getString(8);
                if (rs.getObject(12) != null) {
                qry = qry + rs.getString(12);
                
            }
            }
            be_Consulta.setQuery(qry);
            LBE_BDatos be_BDatos = new LBE_BDatos();
            LBE_Driver driver = new LBE_Driver();
            driver.setDriver(rs.getString(3));
            be_BDatos.setDriver(driver);
            be_BDatos.setUrl(rs.getString(4));
            be_BDatos.setUsuario(rs.getString(5));
            be_BDatos.setClave(rs.getString(6));
            be_BDatos.setCodigo(rs.getInt("id_bdt"));
            be_Consulta.setBe_BDatos(be_BDatos);

            LBE_Grupo Grupo = new LBE_Grupo();
            Grupo.setId(rs.getInt("id_gru"));
            be_Consulta.setGrupo(Grupo);
            be_Consulta.setIdOpcionSegur(rs.getInt("tag_opc"));
            be_Consulta.setFetchSize(rs.getInt("fetch_size"));
        }
        return be_Consulta;
    }

    public void grabaLog(int codConsulta, String codUser, boolean estado,
            String menError) throws SQLException {
        String query = "insert into st_cns_log (id_cns, usr_cns, fec_cns, flg_eje, mns_err) values (?,?,sysdate,?,?)";
        PreparedStatement pst = cn.prepareStatement(query);
        pst.setInt(1, codConsulta);
        pst.setString(2, codUser);
        if (estado) {
            pst.setString(3, "S");
        } else {
            pst.setString(3, "N");
        }
        pst.setString(4, menError);
        pst.executeUpdate();
    }

    public void add(LBE_Consulta consulta, String userLogin) throws SQLException {
        String comando =
                "insert into st_cns\n"
                + "  (id_bdt,\n"
                + "   id_gru,\n"
                + "   id_cns,\n"
                + "   nom_cns,\n"
                + "   qry_cns,\n"
                + "   fec_cre,\n"
                + "   usr_cre,\n"
                + "   tag_opc,\n"
                + "   qry_cns_det1)\n"
                + "values\n"
                + "  (?, ?, ?, ?, ?, sysdate, ?, ?, ?)";
        String selSeq = "select st_sq_cns.nextval from dual";
        consulta.setId(getSeq(selSeq));
        PreparedStatement pst = cn.prepareCall(comando);
        pst.setInt(1, consulta.getBe_BDatos().getCodigo());
        pst.setInt(2, consulta.getGrupo().getId());
        pst.setInt(3, consulta.getId());
        pst.setString(4, consulta.getTitulo());
        String query1 = consulta.getQuery();
        String query2 = "";
        if (consulta.getQuery().length() > 4000) {
            query1 = consulta.getQuery().substring(0, 4000);
            query2 = consulta.getQuery().substring(4000);
        }
        pst.setString(5, query1);
        pst.setString(6, userLogin);
        pst.setInt(7, consulta.getIdOpcionSegur());
        pst.setString(8, query2);
        pst.executeUpdate();
    }

    public void save(LBE_Consulta consulta, String userLogin) throws SQLException {
        String comando =
                "update st_cns set\n"
                + "  id_bdt=?,\n"
                + "   id_gru=?,\n"
                + "   nom_cns=?,\n"
                + "   qry_cns=?,\n"
                + "   fec_mod=sysdate,\n"
                + "   usr_mod=?,\n"
                + "   tag_opc=?,\n"
                + "   qry_cns_det1=?,\n"
                + "   qry_cns_det2=?\n"
                + " where id_cns=?";
        PreparedStatement pst = cn.prepareCall(comando);
        pst.setInt(1, consulta.getBe_BDatos().getCodigo());
        pst.setInt(2, consulta.getGrupo().getId());
        pst.setString(3, consulta.getTitulo());
        String query1 = consulta.getQuery();
        String query2 = "";
        String query3 = "";
        if (consulta.getQuery().length() > 4000) {
            query1 = consulta.getQuery().substring(0, 4000);
            if (consulta.getQuery().length() > 8000) {
                query2 = consulta.getQuery().substring(4000,8000);
                query3 = consulta.getQuery().substring(8000);
            } else {
                query2 = consulta.getQuery().substring(4000);
            }
        }
        pst.setString(4, query1);
        pst.setString(5, userLogin);
        pst.setInt(6, consulta.getIdOpcionSegur());
        pst.setString(7, query2);
        pst.setString(8, query3);
        pst.setInt(9, consulta.getId());
        pst.executeUpdate();
    }

    public ArrayList<LBE_Consulta> getLastByUsr(String usrLogin) throws SQLException {
        ArrayList<LBE_Consulta> elems = new ArrayList<LBE_Consulta>();
        String query =
                "select a1.id_cns, g.nom_gru || '-\\-' || c.nom_cns nom_cons, "
                + " c.tag_opc id_seg, a1.fec_cns last_fecha\n"
                + "from (\n"
                + "select a.id_cns id_cns,\n"
                + "   max(a.fec_cns) fec_cns\n"
                + "from st_cns_log a where Upper(a.usr_cns) = ?\n"
                + "group by a.id_cns\n"
                + "order by fec_cns desc) a1 , st_cns c, st_gru g\n"
                + "where rownum<=10 and a1.id_cns=c.id_cns and c.id_gru=g.id_gru";

        PreparedStatement pst = this.cn.prepareStatement(query);
        pst.setString(1, usrLogin.toUpperCase());
        ResultSet rs = pst.executeQuery();
        int contador = 0;
        while (rs.next()) {
            LBE_Consulta elem = new LBE_Consulta();
            elem.setId(rs.getInt("id_cns"));
            elem.setTitulo(rs.getString("nom_cons"));
            elem.setIdOpcionSegur(rs.getInt("id_seg"));
//            elem.setLastAcceso(rs.getDate("last_fecha") rs.getTime("last_fecha"));
            elems.add(elem);
            contador++;
            if (contador == 10) {
                break;
            }
        }
        return elems;
    }
}
