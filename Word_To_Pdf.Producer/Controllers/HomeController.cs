using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using RabbitMQ.Client;
using System.Diagnostics;
using System.Text;
using Word_To_Pdf.Producer.Models;

namespace Word_To_Pdf.Producer.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IConfiguration _configuration;


        public HomeController(ILogger<HomeController> logger, IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration;
        }

        public IActionResult Index()
        {
            return View();
        }
        public IActionResult WordToPdfPage()
        {
            return View();
        }
        [HttpPost]
        public IActionResult WordToPdfPage(WordToPdf wordToPdf)
        {
            var factory = new ConnectionFactory();
            factory.Uri = new Uri(_configuration["ConnectionStrings:RabbitMQ"]);

            using (var connection = factory.CreateConnection())
            {
                using (var channel = connection.CreateModel())
                {
                    channel.ExchangeDeclare("convert-exchange", ExchangeType.Direct, true, false, null);

                    channel.QueueDeclare(queue: "File", durable: true, exclusive: false, autoDelete: false, arguments: null);


                    channel.QueueBind("File", "convert-exchange", "WordToPdf");

                    MessageWordToPdf messageWordToPdf = new();
                    using (MemoryStream ms = new())
                    {
                        wordToPdf.WordFile.CopyTo(ms);
                        messageWordToPdf.WordByte = ms.ToArray();
                    }

                    messageWordToPdf.Email = wordToPdf.Email;
                    messageWordToPdf.FileName = Path.GetFileNameWithoutExtension(wordToPdf.WordFile.FileName);

                    string serializeMessage = JsonConvert.SerializeObject(messageWordToPdf);

                    byte[] byteMessage = Encoding.UTF8.GetBytes(serializeMessage);


                    var properties = channel.CreateBasicProperties();
                    properties.Persistent = true;

                    channel.BasicPublish("convert-exchange", "WordToPdf", properties,byteMessage);

                    ViewBag.result = "Word dosyanız Pdf dosyasına dönüştürüldükten sonra size Email olarak gönderilecektir";

                    return View();

                }
            }



           
        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}