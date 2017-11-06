package cl.majeanis.satelite.util.json;

import java.io.IOException;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonSerializer;
import com.fasterxml.jackson.databind.SerializerProvider;

import cl.majeanis.satelite.util.tipo.Encrypted;

public class EncryptedSerializer extends JsonSerializer<Encrypted>
{
    @Override
    public void serialize(Encrypted value, JsonGenerator gen, SerializerProvider serializers)
            throws IOException, JsonProcessingException
    {
        gen.writeString(value.text());
    }
}
