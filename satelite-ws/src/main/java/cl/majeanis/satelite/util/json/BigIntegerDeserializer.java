package cl.majeanis.satelite.util.json;

import java.io.IOException;
import java.math.BigInteger;

import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.DeserializationContext;
import com.fasterxml.jackson.databind.JsonDeserializer;

public class BigIntegerDeserializer extends JsonDeserializer<BigInteger>
{
    @Override
    public BigInteger deserialize(JsonParser p, DeserializationContext ctxt) 
            throws IOException, JsonProcessingException
    {
        String s = p.getText().trim();
        if( s.isEmpty() )
            return null;

        return new BigInteger(s);
    }
}
