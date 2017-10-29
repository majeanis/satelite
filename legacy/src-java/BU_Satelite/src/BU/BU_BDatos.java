/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package BU;

import BE.BE_BDatos;
import DB.DB_BDatos;
import Ex.BU_Exception;
import java.sql.SQLException;
import java.util.ArrayList;

/**
 *
 * @author mzavaleta
 */
public class BU_BDatos extends BU_Base{

    public BU_BDatos(BE_BDatos con) throws BU_Exception {
        super(con);
    }
    public ArrayList<BE_BDatos> getAll() throws BU_Exception
    {
        try
        {
            DB_BDatos objBD = new DB_BDatos(getCn(this.con),props);
            return objBD.getAll();
        }catch(SQLException ex)
        {
            throw new BU_Exception("Error al obtener base de datos: " + ex.getMessage());
        }
    }
    public BE_BDatos get(int codigo) throws BU_Exception{
        try
        {
            DB_BDatos objBD = new DB_BDatos(getCn(this.con),props);
            return objBD.get(codigo);
        }catch(SQLException ex)
        {
            throw new BU_Exception("Error al obtener base de datos: " + ex.getMessage());
        }

    }
    public void add(BE_BDatos bDatos, String usrLogin) throws BU_Exception
    {
        try
        {
            DB_BDatos objBD = new DB_BDatos(getCn(this.con),props);
            objBD.add(bDatos, usrLogin);
        }catch(SQLException ex)
        {
            throw new BU_Exception("Error al crear base de datos: " + ex.getMessage());
        }
    }
    public void save(BE_BDatos bDatos, String usrLogin) throws BU_Exception
    {
        try
        {
            DB_BDatos objBD = new DB_BDatos(getCn(this.con),props);
            objBD.save(bDatos, usrLogin);
        }catch(SQLException ex)
        {
            throw new BU_Exception("Error al guardae base de datos: " + ex.getMessage());
        }
    }

}
