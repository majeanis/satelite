package cl.majeanis.satelite.util.po;

import java.sql.CallableStatement;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

import org.apache.ibatis.type.BaseTypeHandler;
import org.apache.ibatis.type.JdbcType;
import org.apache.ibatis.type.MappedJdbcTypes;
import org.apache.ibatis.type.MappedTypes;

import cl.majeanis.satelite.util.tipo.TipoAyuda;

@MappedJdbcTypes(JdbcType.VARCHAR)
@MappedTypes(TipoAyuda.class)
public class TipoAyudaTypeHandler extends BaseTypeHandler<TipoAyuda>
{
    @Override
    public void setNonNullParameter(PreparedStatement ps, int i, TipoAyuda parameter, JdbcType jt) throws SQLException
    {
        ps.setString(i, parameter.name() );
    }

    @Override
    public TipoAyuda getNullableResult(ResultSet rs, String columnName) throws SQLException
    {
        return TipoAyuda.valueOf(rs.getString(columnName));
    }

    @Override
    public TipoAyuda getNullableResult(ResultSet rs, int i) throws SQLException
    {
        return TipoAyuda.valueOf(rs.getString(i));
    }

    @Override
    public TipoAyuda getNullableResult(CallableStatement cs, int i) throws SQLException
    {
        return TipoAyuda.valueOf(cs.getString(i));
    }
}
