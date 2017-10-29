/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package BU;

import BE.BE_BDatos;
import Ex.BU_Exception;
import cripto.Crypto;
import cripto.CryptoException;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.Properties;

/**
 *
 * @author mzavaleta
 */
public class BU_Base {
    protected Properties props = new Properties();
    protected BE_BDatos con;
    protected String getProp(String clave) {
        String nombre = props.getProperty("db");
        return props.getProperty(nombre + "." + clave,"");
    }
    protected Connection getCn() throws BU_Exception
    {
        InputStream is = null;
        String driver="";
        String url="";
        try
        {
            Crypto crypto = new Crypto("MENSAJE DE PAZ");
            is = new FileInputStream("config.properties");
            //is = getClass().getResourceAsStream("config.properties");
            props.load(is);
            is.close();
            String nombre = props.getProperty("db");
            driver =props.getProperty(nombre + ".driver","");
            Class.forName(driver);
            url = props.getProperty(nombre + ".url");
            String user = crypto.desencripta( props.getProperty(nombre + ".user"));
            String password = crypto.desencripta(props.getProperty(nombre + ".password"));
            Connection cn = DriverManager.getConnection(url,user,password);
            return cn;
        }catch(FileNotFoundException ex)
        {
            throw  new BU_Exception("No se encontró archivo de configuracion");
        }
        catch(IOException ex)
        {
            throw  new BU_Exception("Error al leer archivo de configuracion");
        }catch(ClassNotFoundException ex)
        {
            throw  new BU_Exception("Driver: " + driver + " no encontrado");
        }catch(SQLException ex)
        {
            throw new BU_Exception("Error al conectar a Base de datos: " + url);
        }catch(CryptoException ex)
        {
            throw new BU_Exception("Error al desencriptar usuario/clave: " +
                    ex.getMessage());
        }
    }
    protected Connection getCn(BE_BDatos con) throws BU_Exception
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
            throw  new BU_Exception("Driver: " + driver + " no encontrado");
        }catch(SQLException ex)
        {
            throw new BU_Exception("Error al conectar a Base de datos: " + url);
        }
    }
    public BU_Base() throws BU_Exception
    {
        InputStream is = null;
        try
        {
            is = new FileInputStream("config.properties");
            props.load(is);
            is.close();
        }catch(FileNotFoundException ex)
        {
            throw  new BU_Exception("No se encontró archivo de configuracion");
        }
        catch(IOException ex)
        {
            throw  new BU_Exception("Error al leer archivo de configuracion");
        }

    }
    public BU_Base(BE_BDatos con) throws BU_Exception
    {
        this();
        this.con= con;
    }
}
