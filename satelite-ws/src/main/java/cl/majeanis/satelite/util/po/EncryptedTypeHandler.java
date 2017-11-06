package cl.majeanis.satelite.util.po;

import java.sql.CallableStatement;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

import org.apache.ibatis.type.BaseTypeHandler;
import org.apache.ibatis.type.JdbcType;
import org.apache.ibatis.type.MappedJdbcTypes;
import org.apache.ibatis.type.MappedTypes;

import cl.majeanis.satelite.util.tipo.Encrypted;

@MappedJdbcTypes(JdbcType.VARCHAR)
@MappedTypes(Encrypted.class)
public class EncryptedTypeHandler extends BaseTypeHandler<Encrypted>
{
    @Override
    public void setNonNullParameter(PreparedStatement ps, int i, Encrypted parameter, JdbcType jt) throws SQLException
    {
        ps.setString(i, parameter.text() );
    }

    @Override
    public Encrypted getNullableResult(ResultSet rs, String columnName) throws SQLException
    {
        return Encrypted.valueOf(rs.getString(columnName));
    }

    @Override
    public Encrypted getNullableResult(ResultSet rs, int i) throws SQLException
    {
        return Encrypted.valueOf(rs.getString(i));
    }

    @Override
    public Encrypted getNullableResult(CallableStatement cs, int i) throws SQLException
    {
        return Encrypted.valueOf(cs.getString(i));
    }
}
