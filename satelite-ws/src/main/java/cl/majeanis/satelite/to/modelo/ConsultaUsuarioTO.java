package cl.majeanis.satelite.to.modelo;

import cl.majeanis.satelite.to.BaseTO;

public class ConsultaUsuarioTO extends BaseTO
{
    private static final long serialVersionUID = 1L;
    
    private ConsultaTO consulta;
    private UsuarioTO usuario;
    
    public ConsultaTO getConsulta()
    {
        return consulta;
    }
    public void setConsulta(ConsultaTO consulta)
    {
        this.consulta = consulta;
    }
    public UsuarioTO getUsuario()
    {
        return usuario;
    }
    public void setUsuario(UsuarioTO usuario)
    {
        this.usuario = usuario;
    }
}
