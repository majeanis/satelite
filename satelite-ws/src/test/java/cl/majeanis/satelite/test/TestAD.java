package cl.majeanis.satelite.test;

import javax.naming.NamingException;
import javax.naming.ldap.LdapContext;

import cl.majeanis.satelite.util.LoginUtils;

public class TestAD
{

    public static void main(String[] args) throws NamingException
    {
        LdapContext ctx = LoginUtils.loginToAd("luis.millan", "LaClave123", "LOCAL_DS.CL", "192.168.56.14");
        System.out.println( ctx );
        // TODO Auto-generated method stub

    }

}
