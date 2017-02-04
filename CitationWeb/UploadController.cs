using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace CitationWeb
{
    [RoutePrefix("api/upload")]
    public class UploadController : ApiController
    {
        [HttpPost]
        [Route("reciveJson")]
        public IHttpActionResult reciveJson(jsonResult json)
        {
          string path=  System.Web.Hosting.HostingEnvironment.MapPath("~/jsonSampleV4.json");
            File.WriteAllText(path, json.json);
            return Ok();
        }

        //[HttpGet]
        //[Route("reciveJsonGet")]
        //public IHttpActionResult reciveJsonGet(string json) { return Ok("done");
        //}

    }
   public class jsonResult
    {
        public string json { get; set; }

    }
}