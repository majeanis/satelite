/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package DB;

import BE.BE_BDatos;
import BE.BE_Consulta;
import BE.BE_Grupo;
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
public class DB_Grupo extends DB_Base {

    public DB_Grupo(Connection cn) {
        super(cn);
    }

    public DB_Grupo(Connection cn, Properties props) {
        super(cn, props);
    }

    public ArrayList<BE_Grupo> getAllwithCons() throws SQLException {
        ArrayList<BE_Grupo> elems = new ArrayList<BE_Grupo>();
        //String query = "select * from consXGrupo";
        String query = getQuery("grupos");
        PreparedStatement pst = this.cn.prepareStatement(query);
        ResultSet rs = pst.executeQuery();
        BE_Grupo elem = new BE_Grupo();
        while (rs.next()) {
            int idGrupo = rs.getInt(1);
            if (elem.getId() != idGrupo) {
                elem = new BE_Grupo();
                elem.setId(idGrupo);
                elem.setNombre(rs.getString(2));
                elem.setIdOpcionSegur(rs.getInt("tag_opc_g"));
                elems.add(elem);
            }
            BE_Consulta consulta = new BE_Consulta();
            consulta.setId(rs.getInt(3));
            consulta.setTitulo(rs.getString(4));
            consulta.setIdOpcionSegur(rs.getInt("tag_opc_c"));
            elem.getConsultas().add(consulta);
        }
        pst.close();
        return elems;
    }

    public ArrayList<BE_Grupo> getAll() throws SQLException {
        ArrayList<BE_Grupo> elems = new ArrayList<BE_Grupo>();
        String query = "select id_gru, nom_gru, tag_opc from st_gru order by nom_gru";
        //String query = getQuery("grupos");
        PreparedStatement pst = this.cn.prepareStatement(query);
        ResultSet rs = pst.executeQuery();
        while (rs.next()) {
            BE_Grupo elem = new BE_Grupo();
            elem.setId(rs.getInt("id_gru"));
            elem.setNombre(rs.getString("nom_gru"));
            elem.setIdOpcionSegur(rs.getInt("tag_opc"));
            elems.add(elem);
        }
        pst.close();
        return elems;
    }

    public void add(BE_Grupo grupo, String userLogin) throws SQLException {
        String comando =
                "insert into st_gru\n"
                + "  (id_gru, nom_gru, fec_cre, tag_opc,usr_cre)\n"
                + "values\n"
                + "  (ST_SQ_GRU.Nextval,?,sysdate,?,?)";

        PreparedStatement pst = this.cn.prepareStatement(comando);
        pst.setString(1, grupo.getNombre());
        pst.setInt(2, grupo.getIdOpcionSegur());
        pst.setString(3, userLogin);
        pst.executeUpdate();
        pst.close();
    }

    public BE_Grupo get(int codigo) throws SQLException {
        BE_Grupo elem = null;
        String query = "select nom_gru, tag_opc from st_gru where "
                + "id_gru=?";
        PreparedStatement pst = this.cn.prepareStatement(query);
        pst.setInt(1, codigo);
        ResultSet rs = pst.executeQuery();
        if (rs.next()) {
            elem = new BE_Grupo();
            elem.setId(codigo);
            elem.setNombre(rs.getString("nom_gru"));
            elem.setIdOpcionSegur(rs.getInt("tag_opc"));
        }
        pst.close();
        return elem;
    }

    public void save(BE_Grupo grupo, String userLogin) throws SQLException {
        String comando =
                "update st_gru set\n"
                + "  nom_gru=?,tag_opc=?, fec_mod=sysdate, usr_mod=? \n"
                + " where id_gru=?";
        PreparedStatement pst = this.cn.prepareStatement(comando);
        pst.setString(1, grupo.getNombre());
        pst.setInt(2, grupo.getIdOpcionSegur());
        pst.setString(3, userLogin);
        pst.setInt(4, grupo.getId());
        pst.executeUpdate();
        pst.close();
    }

    public ArrayList<BE_Consulta> getConsultas(int id) throws SQLException {
        ArrayList<BE_Consulta> elems = new ArrayList<BE_Consulta>();
        String query =
                "select c.id_cns, c.nom_cns, b.nom_bdt\n"
                + "  from st_cns c, st_bdt b\n"
                + " where c.id_bdt = b.id_bdt\n"
                + "   and c.id_gru = ?\n"
                + " order by c.nom_cns";

        PreparedStatement pst = this.cn.prepareStatement(query);
        pst.setInt(1, id);
        ResultSet rs = pst.executeQuery();
        while (rs.next()) {
            BE_Consulta elem = new BE_Consulta();
            elem.setId(rs.getInt("id_cns"));
            elem.setTitulo(rs.getString("nom_cns"));
            BE_BDatos bDatos = new BE_BDatos();
            bDatos.setNombre(rs.getString("nom_bdt"));
            elem.setBe_BDatos(bDatos);
            elems.add(elem);
        }
        pst.close();
        return elems;
    }
}
