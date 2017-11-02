package cl.majeanis.satelite.util.po;

import java.sql.CallableStatement;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

import org.apache.ibatis.type.BaseTypeHandler;
import org.apache.ibatis.type.JdbcType;
import org.apache.ibatis.type.MappedJdbcTypes;
import org.apache.ibatis.type.MappedTypes;

import cl.majeanis.satelite.util.Utils;
import cl.majeanis.satelite.util.tipo.Horario;

@MappedJdbcTypes(JdbcType.VARCHAR)
@MappedTypes(Horario.class)
public class HorarioTypeHandler extends BaseTypeHandler<Horario>
{
    private static String toJson(Horario horario)
    {
        return Utils.toJson(horario);
    }
    
    private static Horario fromJson(String value)
    {
        return Utils.fromJson(Horario.class, value);
    }
    
    @Override
    public void setNonNullParameter(PreparedStatement ps, int i, Horario parameter, JdbcType jt) throws SQLException
    {
        ps.setString(i, toJson( parameter ) );
    }

    @Override
    public Horario getNullableResult(ResultSet rs, String columnName) throws SQLException
    {
        return fromJson( rs.getString(columnName) );
    }

    @Override
    public Horario getNullableResult(ResultSet rs, int i) throws SQLException
    {
        return fromJson( rs.getString(i) );
    }

    @Override
    public Horario getNullableResult(CallableStatement cs, int i) throws SQLException
    {
        return fromJson( cs.getString(i) );
    }
}
