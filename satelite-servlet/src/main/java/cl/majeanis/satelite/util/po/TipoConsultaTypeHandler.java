package cl.majeanis.satelite.util.po;

import java.sql.CallableStatement;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

import org.apache.ibatis.type.BaseTypeHandler;
import org.apache.ibatis.type.JdbcType;
import org.apache.ibatis.type.MappedJdbcTypes;
import org.apache.ibatis.type.MappedTypes;

import cl.majeanis.satelite.util.tipo.TipoConsulta;

@MappedJdbcTypes(JdbcType.VARCHAR)
@MappedTypes(TipoConsulta.class)
public class TipoConsultaTypeHandler extends BaseTypeHandler<TipoConsulta>
{
    @Override
    public void setNonNullParameter(PreparedStatement ps, int i, TipoConsulta parameter, JdbcType jt) throws SQLException
    {
        ps.setString(i, parameter.name() );
    }

    @Override
    public TipoConsulta getNullableResult(ResultSet rs, String columnName) throws SQLException
    {
        return TipoConsulta.valueOf(rs.getString(columnName));
    }

    @Override
    public TipoConsulta getNullableResult(ResultSet rs, int i) throws SQLException
    {
        return TipoConsulta.valueOf(rs.getString(i));
    }

    @Override
    public TipoConsulta getNullableResult(CallableStatement cs, int i) throws SQLException
    {
        return TipoConsulta.valueOf(cs.getString(i));
    }
}
