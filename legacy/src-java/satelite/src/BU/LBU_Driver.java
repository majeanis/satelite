/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package BU;


import BD.LDB_Driver;
import BE.LBE_BDatos;
import BE.LBE_Driver;
import BU.Ex.LBU_Exception;
import java.sql.SQLException;
import java.util.ArrayList;

/**
 *
 * @author mzavaleta
 */
public class LBU_Driver extends LBU_Base{
    public LBU_Driver(LBE_BDatos con) throws LBU_Exception {
        super(con);
    }
    public ArrayList<LBE_Driver> getAll() throws LBU_Exception
    {
        try
        {
            LDB_Driver objDB = new LDB_Driver(getCn(this.con),props);
            return objDB.getAll();
        }catch(SQLException ex)
        {
            ex.printStackTrace();
            throw new LBU_Exception("Error al obtener drivers: " + ex.getMessage());
        }
    }


}
