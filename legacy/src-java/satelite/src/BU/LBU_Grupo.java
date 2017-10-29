/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package BU;


import BD.LDB_Grupo;
import BE.LBE_BDatos;
import BE.LBE_Consulta;
import BE.LBE_Grupo;
import BU.Ex.LBU_Exception;
import java.sql.SQLException;
import java.util.ArrayList;

/**
 *
 * @author mzavaleta
 */
public class LBU_Grupo extends LBU_Base{
    public LBU_Grupo(LBE_BDatos con) throws LBU_Exception
    {
        super(con);
    }
    public ArrayList<LBE_Grupo> getAllwithCons() throws LBU_Exception
    {
        try
        {
            LDB_Grupo objDB = new LDB_Grupo(getCn(this.con),props);
            return objDB.getAllwithCons();
        }catch(SQLException ex)
        {
            throw new LBU_Exception("Error al obtener grupos: " + ex.getMessage());
        }
    }
    public ArrayList<LBE_Grupo> getAll() throws LBU_Exception
    {
        try
        {
            LDB_Grupo objDB = new LDB_Grupo(getCn(this.con),props);
            return objDB.getAll();
        }catch(SQLException ex)
        {
            throw new LBU_Exception("Error al obtener grupos: " + ex.getMessage());
        }
    }
    public LBE_Grupo get(int codigo) throws LBU_Exception
    {
        try
        {
            LDB_Grupo objDB = new LDB_Grupo(getCn(this.con),props);
            return objDB.get(codigo);
        }catch(SQLException ex)
        {
            throw new LBU_Exception("Error al obtener grupo: " + ex.getMessage());
        }
    }
    public void add(LBE_Grupo elem, String userLogin) throws LBU_Exception
    {
        try
        {
            LDB_Grupo objDB = new LDB_Grupo(getCn(this.con),props);
            objDB.add(elem, userLogin);
        }catch(SQLException ex)
        {
            throw new LBU_Exception("Error al crear grupo: " + ex.getMessage());
        }
    }
    public void save(LBE_Grupo elem, String userLogin) throws LBU_Exception
    {
        try
        {
            LDB_Grupo objDB = new LDB_Grupo(getCn(this.con),props);
            objDB.save(elem, userLogin);
        }catch(SQLException ex)
        {
            throw new LBU_Exception("Error al guardar grupo: " + ex.getMessage());
        }
    }
    public ArrayList<LBE_Consulta> getConsultas(int id) throws LBU_Exception
    {
        try
        {
            LDB_Grupo objDB = new LDB_Grupo(getCn(this.con),props);
            return objDB.getConsultas(id);
        }catch(SQLException ex)
        {
            throw new LBU_Exception("Error al obtener consultas para grupo : " + ex.getMessage());
        }
    }

}
