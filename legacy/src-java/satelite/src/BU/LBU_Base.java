/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package BU;

import BE.LBE_BDatos;
import BU.Ex.LBU_Exception;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.Properties;
import lcripto.LCrypto;
import lcripto.LCryptoException;
import logger.Log;

/**
 *
 * @author mzavaleta
 */
public class LBU_Base {
    protected Properties props = new Properties();
    protected LBE_BDatos con;
    protected String getProp(String clave) {
        String nombre = props.getProperty("db");
        return props.getProperty(nombre + "." + clave,"");
    }
    protected Connection getCn() throws LBU_Exception
    {
        InputStream is = null;
        String driver="";
        String url="";
        try
        {
            LCrypto crypto = new LCrypto("MENSAJE DE PAZ");
            is = new FileInputStream("config.properties");
            //is = getClass().getResourceAsStream("config.properties");
            props.load(is);
            is.close();
            String nombre = props.getProperty("db");
            driver =props.getProperty(nombre + ".driver","");
            Log.debug("Carga de drive inicio");
            Class.forName(driver);
            Log.debug("Carga de drive fin");
            url = props.getProperty(nombre + ".url");
            Log.debug("Descencripta inicio");
            String user = crypto.desencripta( props.getProperty(nombre + ".user"));
            String password = crypto.desencripta(props.getProperty(nombre + ".password"));
            Log.debug("Desencripta fin");
            Log.debug("GetConn inicio");
            Connection cn = DriverManager.getConnection(url,user,password);
            Log.debug("Carga de drive fin");
            return cn;
        }catch(FileNotFoundException ex)
        {
            throw  new LBU_Exception("No se encontró archivo de configuracion");
        }
        catch(IOException ex)
        {
            throw  new LBU_Exception("Error al leer archivo de configuracion");
        }catch(ClassNotFoundException ex)
        {
            throw  new LBU_Exception("Driver: " + driver + " no encontrado");
        }catch(SQLException ex)
        {
            throw new LBU_Exception("Error al conectar a Base de datos: " + url);
        }catch(LCryptoException ex)
        {
            throw new LBU_Exception("Error al desencriptar usuario/clave: " +
                    ex.getMessage());
        }
    }
    protected Connection getCn(LBE_BDatos con) throws LBU_Exception
    {
        String driver="";
        String url="";
        try
        {
            driver =con.getDriver().getDriver();
            Class.forName(driver);
            String user = con.getUsuario();
            String password = con.getClave();
            url =  con.getUrl();
            Connection cn = DriverManager.getConnection(url,user,password);
            return cn;
        }catch(ClassNotFoundException ex)
        {
            throw  new LBU_Exception("Driver: " + driver + " no encontrado");
        }catch(SQLException ex)
        {
            throw new LBU_Exception("Error al conectar a Base de datos: " + url);
        }
    }
    public LBU_Base() throws LBU_Exception
    {
        InputStream is = null;
        try
        {
            is = new FileInputStream("config.properties");
            props.load(is);
            is.close();
        }catch(FileNotFoundException ex)
        {
            throw  new LBU_Exception("No se encontró archivo de configuracion");
        }
        catch(IOException ex)
        {
            throw  new LBU_Exception("Error al leer archivo de configuracion");
        }

    }
    public LBU_Base(LBE_BDatos con) throws LBU_Exception
    {
        this();
        this.con= con;
    }
}
