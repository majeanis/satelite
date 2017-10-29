/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package BU;

import BD.LDB_BDatos;
import BE.LBE_BDatos;
import BU.Ex.LBU_Exception;
import java.sql.SQLException;
import java.util.ArrayList;

/**
 *
 * @author mzavaleta
 */
public class LBU_BDatos extends LBU_Base{

    public LBU_BDatos(LBE_BDatos con) throws LBU_Exception {
        super(con);
    }
    public ArrayList<LBE_BDatos> getAll() throws LBU_Exception
    {
        try
        {
            LDB_BDatos objBD = new LDB_BDatos(getCn(this.con),props);
            return objBD.getAll();
        }catch(SQLException ex)
        {
            throw new LBU_Exception("Error al obtener base de datos: " + ex.getMessage());
        }
    }
    public LBE_BDatos get(int codigo) throws LBU_Exception{
        try
        {
            LDB_BDatos objBD = new LDB_BDatos(getCn(this.con),props);
            return objBD.get(codigo);
        }catch(SQLException ex)
        {
            throw new LBU_Exception("Error al obtener base de datos: " + ex.getMessage());
        }

    }
    public void add(LBE_BDatos bDatos, String usrLogin) throws LBU_Exception
    {
        try
        {
            LDB_BDatos objBD = new LDB_BDatos(getCn(this.con),props);
            objBD.add(bDatos, usrLogin);
        }catch(SQLException ex)
        {
            throw new LBU_Exception("Error al crear base de datos: " + ex.getMessage());
        }
    }
    public void save(LBE_BDatos bDatos, String usrLogin) throws LBU_Exception
    {
        try
        {
            LDB_BDatos objBD = new LDB_BDatos(getCn(this.con),props);
            objBD.save(bDatos, usrLogin);
        }catch(SQLException ex)
        {
            throw new LBU_Exception("Error al guardae base de datos: " + ex.getMessage());
        }
    }

}
