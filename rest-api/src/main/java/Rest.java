import lombok.val;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.Response;

import java.io.IOException;

public class Rest {

    private static OkHttpClient client = new OkHttpClient();

    private static String run() throws IOException {
        val request = new Request
                .Builder()
                .url("http://vk.com/")
                .build();

        Response response = client.newCall(request).execute();
        return response.body().string();
    }

    static public void main(String... args) throws IOException {
        System.out.println(run());
    }
}
