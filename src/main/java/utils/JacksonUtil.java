package utils;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.ObjectMapper;


public class JacksonUtil {

    public static <R> Object getResponseAsObject(String response, Class<R> classType) {
        ObjectMapper mapper = new ObjectMapper();
        mapper.configure(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES, false);

        try {
            return mapper.readValue(response, classType);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }
    public static String getAsString(Object obj){
        ObjectMapper objectMapper = new ObjectMapper();
        String asString = "";
        try {
            asString = objectMapper.writeValueAsString(obj);
        } catch (JsonProcessingException e) {
            e.printStackTrace();
        }
        return asString;
    }
}
