/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package BU;

import BE.BE_BDatos;
import BE.BE_Consulta;
import BE.BE_Grupo;
import DB.DB_Grupo;
import Ex.BU_Exception;
import java.sql.SQLException;
import java.util.ArrayList;

/**
 *
 * @author mzavaleta
 */
public class BU_Grupo extends BU_Base{
    public BU_Grupo(BE_BDatos con) throws BU_Exception
    {
        super(con);
    }
    public ArrayList<BE_Grupo> getAllwithCons() throws BU_Exception
    {
        try
        {
            DB_Grupo objDB = new DB_Grupo(getCn(this.con),props);
            return objDB.getAllwithCons();
        }catch(SQLException ex)
        {
            throw new BU_Exception("Error al obtener grupos: " + ex.getMessage());
        }
    }
    public ArrayList<BE_Grupo> getAll() throws BU_Exception
    {
        try
        {
            DB_Grupo objDB = new DB_Grupo(getCn(this.con),props);
            return objDB.getAll();
        }catch(SQLException ex)
        {
            throw new BU_Exception("Error al obtener grupos: " + ex.getMessage());
        }
    }
    public BE_Grupo get(int codigo) throws BU_Exception
    {
        try
        {
            DB_Grupo objDB = new DB_Grupo(getCn(this.con),props);
            return objDB.get(codigo);
        }catch(SQLException ex)
        {
            throw new BU_Exception("Error al obtener grupo: " + ex.getMessage());
        }
    }
    public void add(BE_Grupo elem, String userLogin) throws BU_Exception
    {
        try
        {
            DB_Grupo objDB = new DB_Grupo(getCn(this.con),props);
            objDB.add(elem, userLogin);
        }catch(SQLException ex)
        {
            throw new BU_Exception("Error al crear grupo: " + ex.getMessage());
        }
    }
    public void save(BE_Grupo elem, String userLogin) throws BU_Exception
    {
        try
        {
            DB_Grupo objDB = new DB_Grupo(getCn(this.con),props);
            objDB.save(elem, userLogin);
        }catch(SQLException ex)
        {
            throw new BU_Exception("Error al guardar grupo: " + ex.getMessage());
        }
    }
    public ArrayList<BE_Consulta> getConsultas(int id) throws BU_Exception
    {
        try
        {
            DB_Grupo objDB = new DB_Grupo(getCn(this.con),props);
            return objDB.getConsultas(id);
        }catch(SQLException ex)
        {
            throw new BU_Exception("Error al obtener consultas para grupo : " + ex.getMessage());
        }
    }

}
