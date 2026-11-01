package id.irvan.kantorvanx;

import android.annotation.SuppressLint;
import android.app.Activity;
import android.content.Intent;
import android.graphics.Color;
import android.net.Uri;
import android.os.Bundle;
import android.webkit.WebResourceError;
import android.webkit.WebResourceRequest;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.Toast;

// Kantor VanX — aplikasi web-to-app: membungkus halaman perintah jarak jauh
// Kantor VanX (https://irvanmaulana.my.id/kantor/) dalam WebView.
// Halaman menyimpan token GitHub user di localStorage (DomStorage), jadi
// token tetap tersimpan antar sesi aplikasi.
public class MainActivity extends Activity {
    private WebView webView;
    private boolean showingError;

    private static final String ERROR_HTML =
            "<html><head><meta name='viewport' content='width=device-width, initial-scale=1'>" +
            "<style>body{background:#0d1117;color:#e6edf3;font-family:sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;text-align:center}" +
            "a{display:inline-block;margin-top:18px;background:#f97316;color:#14100c;padding:14px 28px;border-radius:12px;text-decoration:none;font-weight:bold}</style></head>" +
            "<body><div><div style='font-size:44px'>&#128225;</div><h2>Koneksi terputus</h2>" +
            "<p>Kantor VanX butuh internet untuk menyala.</p>" +
            "<a href='" + BuildConfig.WEB_APP_URL + "'>COBA LAGI</a></div></body></html>";

    @SuppressLint("SetJavaScriptEnabled")
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        try {
            getWindow().setStatusBarColor(Color.rgb(13, 17, 23));
            getWindow().setNavigationBarColor(Color.rgb(13, 17, 23));

            webView = new WebView(this);
            webView.setBackgroundColor(Color.rgb(13, 17, 23));
            setContentView(webView);

            WebSettings settings = webView.getSettings();
            settings.setJavaScriptEnabled(true);
            settings.setDomStorageEnabled(true);
            settings.setDatabaseEnabled(true);
            settings.setAllowFileAccess(false);
            settings.setAllowContentAccess(true);
            settings.setMixedContentMode(WebSettings.MIXED_CONTENT_NEVER_ALLOW);
            settings.setUserAgentString(settings.getUserAgentString() + " KantorVanXAndroid/1.0");

            webView.setWebViewClient(new WebViewClient() {
                @Override
                public boolean shouldOverrideUrlLoading(WebView view, WebResourceRequest request) {
                    Uri uri = request.getUrl();
                    String host = uri.getHost();
                    if (host != null && (host.equals("irvanmaulana.my.id") || host.equals("github.com"))) return false;
                    try {
                        startActivity(new Intent(Intent.ACTION_VIEW, uri));
                        return true;
                    } catch (Exception ignored) {
                        return false;
                    }
                }

                @Override
                public void onPageFinished(WebView view, String url) {
                    if (url.startsWith("https://")) showingError = false;
                }

                @Override
                public void onReceivedError(WebView view, WebResourceRequest request, WebResourceError error) {
                    if (request.isForMainFrame() && !showingError) {
                        showingError = true;
                        view.loadDataWithBaseURL(null, ERROR_HTML, "text/html", "UTF-8", null);
                    }
                }
            });

            webView.loadUrl(BuildConfig.WEB_APP_URL);
        } catch (Exception e) {
            e.printStackTrace();
            try {
                Toast.makeText(this, "Error: " + e.getMessage(), Toast.LENGTH_LONG).show();
            } catch (Exception ignored) {}
        }
    }

    @Override
    public void onBackPressed() {
        try {
            if (webView != null && webView.canGoBack()) {
                webView.goBack();
            } else {
                super.onBackPressed();
            }
        } catch (Exception e) {
            super.onBackPressed();
        }
    }

    @Override
    protected void onDestroy() {
        try {
            if (webView != null) webView.destroy();
        } catch (Exception ignored) {}
        super.onDestroy();
    }
}
