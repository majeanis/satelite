package cl.majeanis.satelite.util.json;

import java.io.File;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Writer;
import java.math.BigInteger;
import java.text.SimpleDateFormat;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import com.fasterxml.jackson.annotation.JsonAutoDetect;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.PropertyAccessor;
import com.fasterxml.jackson.core.JsonGenerator.Feature;
import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.databind.deser.std.NumberDeserializers.BigIntegerDeserializer;
import com.fasterxml.jackson.databind.module.SimpleModule;
import com.fasterxml.jackson.datatype.jdk8.Jdk8Module;
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule;

import cl.majeanis.satelite.util.tipo.Encrypted;

public class JacksonFacade
{
    private static final Logger logger = LogManager.getLogger(JacksonFacade.class);

    private static final ObjectMapper mapper = newMapper();

    private static ObjectMapper newMapper()
    {
        ObjectMapper mapper = new ObjectMapper();

        mapper.disable(SerializationFeature.FAIL_ON_EMPTY_BEANS);
        mapper.disable(SerializationFeature.WRITE_DATES_AS_TIMESTAMPS);

        mapper.enable(Feature.IGNORE_UNKNOWN);
        mapper.enable(Feature.WRITE_BIGDECIMAL_AS_PLAIN);
   
        mapper.enable(SerializationFeature.INDENT_OUTPUT);
        mapper.enable(DeserializationFeature.USE_BIG_INTEGER_FOR_INTS);
        mapper.enable(DeserializationFeature.USE_BIG_DECIMAL_FOR_FLOATS);
        
        mapper.setDateFormat(new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSS"));

        /**
         * Registramos los módulos mínimos necesarios para el formateado de los objetos 
         */
        SimpleModule module = new SimpleModule();
        module.addSerializer(BigInteger.class, new BigIntegerSerializer() );
        module.addDeserializer(BigInteger.class, new BigIntegerDeserializer());
        
        module.addSerializer(Encrypted.class, new EncryptedSerializer() );
        module.addDeserializer(Encrypted.class, new EncryptedDeserializer());

        mapper.registerModule(module);
        mapper.registerModule(new JavaTimeModule());
        mapper.registerModule(new Jdk8Module());
        
        mapper.setVisibility(PropertyAccessor.FIELD, JsonAutoDetect.Visibility.ANY);
        mapper.setVisibility(PropertyAccessor.GETTER, JsonAutoDetect.Visibility.NONE);
        mapper.setVisibility(PropertyAccessor.IS_GETTER, JsonAutoDetect.Visibility.NONE);

        mapper.setSerializationInclusion(JsonInclude.Include.NON_NULL);

        return mapper;
    }

    public static String toJson(Object obj)
    {
        try 
        {
            return mapper.writeValueAsString(obj);
        } catch (Exception e) 
        {
            logger.error( "toJson[ERR] al generar json de: obj=" + obj, e);
            return null;
        }
    }

    public static void toJson(Object obj, File file)
    {
        try
        {
            mapper.writeValue(file, obj);
        }
        catch(Exception e)
        {
            logger.error("toJson[ERR] al serializar objeto en un File: obj=" + obj + " file=" + file, e );            
        }
    }

    public static void toJson(Object obj, OutputStream os)
    {
        try
        {
            mapper.writeValue(os, obj);
        }
        catch(Exception e)
        {
            logger.error("toJson[ERR] al serializar objeto en un Stream: obj=" + obj + " os=" + os, e );
        }
    }

    public static void toJson(Object obj, Writer writer)
    {
        try
        {
            mapper.writeValue(writer, obj);
        }
        catch(Exception e)
        {
            logger.error("toJson[ERR] al serializar objeto en un Stream: obj=" + obj + " writer=" + writer, e );
        }
    }

    public static <T> T fromJson(Class<T> clazz, String json)
    {
        try 
        {
            return mapper.readValue(json, clazz);
        } catch (Exception e) 
        {
            logger.error("fromJson[ERR] al recrear objeto desde el String: json=" + json, e );
            return null;
        }
    }

    public static <T> T fromJson(Class<T> clazz, InputStream is)
    {
        try 
        {
            return mapper.readValue(is, clazz);
        } catch (Exception e) 
        {
            logger.error("fromJson[ERR] al recrear objeto desde un InputStream:", e );          
            return null;
        }
    }

    public static <T> T fromJson(Class<T> clazz, File src)
    {
        try 
        {
            return mapper.readValue(src, clazz);
        } catch (Exception e) 
        {
            logger.error("fromJson[ERR] al recrear objeto desde unn File: src=" + src, e );            
            return null;
        }
    }
}
