using System;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace Reporte_costos_xtarima
{
    /// <summary>
    /// Envia correo por Microsoft Graph (POST /users/{buzon}/sendMail) con credenciales de aplicacion.
    /// Sin dependencias externas: el ejecutable es autocontenido (el secreto va en GraphSecreto.cs).
    /// </summary>
    public static class GraphMail
    {
        private const string ClientId = "c982ac22-b59a-441c-905a-9eb273ade26b";
        private const string TenantId = "d20460dd-675d-4b51-87cc-9d10f9175633";

        public static void Enviar(string buzonRemitente, string para, string asunto, string cuerpoHtml)
        {
            // .NET 4.0 no define Tls12 en el enum; Microsoft exige TLS 1.2.
            ServicePointManager.SecurityProtocol |= (SecurityProtocolType)3072;

            string token = ObtenerToken();

            string json = "{\"message\":{\"subject\":" + Json(asunto) +
                ",\"body\":{\"contentType\":\"HTML\",\"content\":" + Json(cuerpoHtml) + "}" +
                ",\"toRecipients\":[{\"emailAddress\":{\"address\":" + Json(para) + "}}]}" +
                ",\"saveToSentItems\":true}";

            string url = "https://graph.microsoft.com/v1.0/users/" + Uri.EscapeDataString(buzonRemitente) + "/sendMail";
            Post(url, "application/json", json, token);
        }

        private static string ObtenerToken()
        {
            string form = "client_id=" + Uri.EscapeDataString(ClientId) +
                "&client_secret=" + Uri.EscapeDataString(GraphSecreto.Valor) +
                "&scope=" + Uri.EscapeDataString("https://graph.microsoft.com/.default") +
                "&grant_type=client_credentials";

            string resp = Post("https://login.microsoftonline.com/" + TenantId + "/oauth2/v2.0/token",
                "application/x-www-form-urlencoded", form, null);

            Match m = Regex.Match(resp, "\"access_token\"\\s*:\\s*\"([^\"]+)\"");
            if (!m.Success)
                throw new InvalidOperationException("Respuesta de token sin access_token: " + resp);
            return m.Groups[1].Value;
        }

        private static string Post(string url, string contentType, string body, string bearer)
        {
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = "POST";
            req.ContentType = contentType;
            if (bearer != null)
                req.Headers["Authorization"] = "Bearer " + bearer;

            byte[] datos = Encoding.UTF8.GetBytes(body);
            req.ContentLength = datos.Length;
            using (Stream s = req.GetRequestStream())
                s.Write(datos, 0, datos.Length);

            try
            {
                using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                using (StreamReader sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
                    return sr.ReadToEnd();
            }
            catch (WebException ex)
            {
                string detalle = ex.Message;
                if (ex.Response != null)
                    using (StreamReader sr = new StreamReader(ex.Response.GetResponseStream(), Encoding.UTF8))
                        detalle = sr.ReadToEnd();
                throw new InvalidOperationException("Graph respondio error: " + detalle, ex);
            }
        }

        private static string Json(string s)
        {
            StringBuilder sb = new StringBuilder("\"");
            foreach (char c in s)
            {
                switch (c)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (c < 0x20) sb.Append("\\u" + ((int)c).ToString("x4"));
                        else sb.Append(c);
                        break;
                }
            }
            return sb.Append('"').ToString();
        }
    }
}
