/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package BU;

import BE.BE_BDatos;
import BE.BE_Driver;
import DB.DB_Driver;
import Ex.BU_Exception;
import java.sql.SQLException;
import java.util.ArrayList;

/**
 *
 * @author mzavaleta
 */
public class BU_Driver extends BU_Base{
    public BU_Driver(BE_BDatos con) throws BU_Exception {
        super(con);
    }
    public ArrayList<BE_Driver> getAll() throws BU_Exception
    {
        try
        {
            DB_Driver objDB = new DB_Driver(getCn(this.con),props);
            return objDB.getAll();
        }catch(SQLException ex)
        {
            ex.printStackTrace();
            throw new BU_Exception("Error al obtener drivers: " + ex.getMessage());
        }
    }


}
