import java.util.Map;
import java.util.TreeMap;

public class ColorResolver {
    static Map<String, String> colors = new TreeMap<String, String>();

    static {
       colors.put("темно-серый", "Dark Grey");
       colors.put("черный", "Black");
       colors.put("белый", "White");
    }

    public static String resolveColor(String color) {
        String resolvedColor = colors.get(color.toLowerCase());
        if (resolvedColor == null)
            return color;
        return resolvedColor;
    }
}
