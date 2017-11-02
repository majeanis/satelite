package cl.majeanis.satelite.util.json;

import java.io.IOException;

import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.DeserializationContext;
import com.fasterxml.jackson.databind.JsonDeserializer;

import cl.majeanis.satelite.util.tipo.Encrypted;

public class EncryptedDeserializer extends JsonDeserializer<Encrypted>
{
    @Override
    public Encrypted deserialize(JsonParser p, DeserializationContext ctxt) 
            throws IOException, JsonProcessingException
    {
        String s = p.getText().trim();
        return Encrypted.valueOf(s);
    }
}
