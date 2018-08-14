using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using JFetchHome;
using System.Net.Http;
using Newtonsoft.Json;

namespace JFetchExcelHome
{
	public static class ExcelAddin {

		private static bool loggedIn = false;
		private static HttpClient client = new HttpClient();

		[ExcelFunction(Description = "Print Info about Kings")]
		public static object GetKings() {
			return ExcelAsyncUtil.Run("GetKings", new object[] { }, () => JFetch.JFetchSync("https://jsonplaceholder.typicode.com/posts", client));
		}

		/*[ExcelFunction(Description = "Download portfolio data from Orion")]
		public static object Orion_FP_Focus(string shortname, string date) {

		}*/

		private static async void authAsync(string username, string password) {
			client.DefaultRequestHeaders.Add("Authentication", "Basic " + Base64Encode(username + ":" + password));
			var response = await client.GetAsync("http://authurl.url/v1/security/token").ConfigureAwait(false);
			response.EnsureSuccessStatusCode();
			var j = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
			//var d = JsonConvert.DeserializeObject
			if (true) {
				loggedIn = true;
			}
		}

		internal static string Base64Encode(string plaintext) {
			var plaintextbytes = System.Text.Encoding.UTF8.GetBytes(plaintext);
			return System.Convert.ToBase64String(plaintextbytes);
		}
	}
} 