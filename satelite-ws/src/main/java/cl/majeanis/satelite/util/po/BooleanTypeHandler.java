package cl.majeanis.satelite.util.po;

import java.sql.CallableStatement;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

import org.apache.ibatis.type.BaseTypeHandler;
import org.apache.ibatis.type.JdbcType;
import org.apache.ibatis.type.MappedTypes;

@MappedTypes(Boolean.class)
public class BooleanTypeHandler extends BaseTypeHandler<Boolean>
{
    @Override
    public Boolean getNullableResult(ResultSet rs, String columnName) throws SQLException
    {
        String value = rs.getString(columnName);
        return "S".equalsIgnoreCase(value);
    }

    @Override
    public Boolean getNullableResult(ResultSet rs, int columnIndex) throws SQLException
    {
        String value = rs.getString(columnIndex);
        return "S".equalsIgnoreCase(value);        
    }

    @Override
    public Boolean getNullableResult(CallableStatement cs, int columnIndex) throws SQLException
    {
        String value = cs.getString(columnIndex);
        return "S".equalsIgnoreCase(value);        
    }

    @Override
    public void setNonNullParameter(PreparedStatement ps, int i, Boolean parameter, JdbcType jdbcType) throws SQLException
    {
        if (parameter == null)
        {
            ps.setNull(i, java.sql.Types.VARCHAR);
        } else
        {
            ps.setString(i, (parameter ? "S": "N"));
        }
    }
}
