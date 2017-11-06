package cl.majeanis.satelite.util;

import java.util.Hashtable;

import javax.naming.AuthenticationException;
import javax.naming.CommunicationException;
import javax.naming.Context;
import javax.naming.NamingException;
import javax.naming.ldap.InitialLdapContext;
import javax.naming.ldap.LdapContext;

import org.apache.commons.lang3.StringUtils;

public final class LoginUtils
{
    public static LdapContext loginToAd(String username, String password, String domainName, String serverName)
            throws AuthenticationException, CommunicationException, NamingException
    {
        if (StringUtils.isBlank(password))
        {
            password = null;
        }

        if (StringUtils.isBlank(serverName))
        {
            serverName = "127.0.0.1";
        }

        /**
         * Preparación de las propiedades para abrir el ContectoLdap
         */
        Hashtable<String, Object> props = new Hashtable<>();

        String principalName = username + "@" + domainName;
        props.put(Context.SECURITY_PRINCIPAL, principalName);

        if (password != null)
        {
            props.put(Context.SECURITY_CREDENTIALS, password);
        }

        String ldapURL = "ldap://" + ((serverName == null) ? domainName : serverName) + '/';
        props.put(Context.INITIAL_CONTEXT_FACTORY, "com.sun.jndi.ldap.LdapCtxFactory");
        props.put(Context.PROVIDER_URL, ldapURL);

        return new InitialLdapContext(props, null);
    }

    public static LdapContext loginToAd(String username, String password, String domainName) throws NamingException
    {
        return loginToAd(username, password, domainName, null);
    }
}