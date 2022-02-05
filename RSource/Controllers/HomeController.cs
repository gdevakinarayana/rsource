using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using RSource.Models;
using Resouce.ViewModels;

namespace Resouce.ViewModels
{
    public class StudentInfo
    {
        public int Id { get; set; }
        public string RollNumber { get; set; }
        public string StudentName { get; set; }
        public string Skill { get; set; }
    }

    public class HashSetData
    {
        public int Percentage { get; set; }
        public string Skill { get; set; }
    }
}

namespace RSource.Controllers
{

    public static class Utility
    {
        public static List<StudentInfo> MergeTables(DataTable students, DataTable skills)
        {
            var JoinResult = (from p in students.AsEnumerable()
                              join t in skills.AsEnumerable()
                              on p.Field<double>("S.No") equals t.Field<int>("Id")
                              select new StudentInfo
                              {
                                  Id = t.Field<int>("Id"),
                                  RollNumber = p.Field<string>("Roll No."),
                                  StudentName = p.Field<string>("Student Name"),
                                  Skill = t.Field<string>("Skill")
                              }).ToList();
            return JoinResult;
        }

        public static DataTable RandomRows(List<HashSetData> skillslist, int rowcount)
        {
            var schema = GetTableSchema();
            //var length = table.Rows.Count;
            var _numberRowslist = Enumerable.Range(0, rowcount).ToList();
            var random = new Random();
            //var result = (50*30)/100;
            foreach (HashSetData hashSetData in skillslist)
            {
                var percentage = hashSetData.Percentage;
                var _p = (rowcount * percentage / 100);

                Random r = new Random();
                var selectedNumbers = _numberRowslist.OrderBy(x => r.Next()).Take(_p).ToList();
                //var selectedNumbers = _numberRowslist.Take(_p).ToList();

                for (int index = 0; index < _p; index++)
                {
                    var id = selectedNumbers[index];
                    var _newRow = schema.NewRow();
                    _newRow[0] = id;
                    _newRow[1] = hashSetData.Skill;
                    schema.Rows.Add(_newRow);
                    _numberRowslist.Remove(id);
                }
            }
            return schema;
        }
        public static DataTable GetTableSchema()
        {
            var dt = new DataTable();
            dt.Columns.Add(new DataColumn("Id", typeof(int)));
            dt.Columns.Add(new DataColumn("Skill", typeof(string)));
            return dt;
        }

        public static DataTable Read(string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        UseColumnDataType = false,
                        ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                        {

                            UseHeaderRow = true
                        }
                    });
                    return result.Tables[0];
                }
            }
        }
    }

    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment hostingEnvironment;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment environment)
        {
            _logger = logger;
            hostingEnvironment = environment;
        }

        public IActionResult DataRefresh()
        {
            var subDirectory = "upload";
            var _dPath = Path.Combine(hostingEnvironment.ContentRootPath, subDirectory);
            var _filePath = "";
            if (Directory.Exists(_dPath))
            {
                var files = Directory.GetFiles(_dPath, "*.xls*");
                if (files.Length > 0)
                    _filePath = files.First();
            }
            if (string.IsNullOrEmpty(_filePath))
            {
                return View();
            }
            var studenDataSet = Utility.Read(_filePath);
            var _r = studenDataSet.Rows.Count;

            var skills = getSkills();
            var skillDataSet = Utility.RandomRows(skills, _r);

            var _mset = Utility.MergeTables(studenDataSet, skillDataSet);

            ViewBag.Data = _mset;
            return View("Index");
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> UploadFile()
        {
            var files = Request.Form.Files;
            IFormFile formFile=null;
            if (files.Count() > 0)
            {
                formFile = files[0];
            }
            if (formFile == null)
                View("Index");
            var _filePath = await SaveFile(formFile);

            var studenDataSet = Utility.Read(_filePath);
            var _r = studenDataSet.Rows.Count;

            var skills = getSkills();
            var skillDataSet = Utility.RandomRows(skills, _r);

            var _mset = Utility.MergeTables(studenDataSet, skillDataSet);

            ViewBag.Data = _mset;
            return View("Index");
        }

        private List<HashSetData> getSkills()
        {
            var list = new List<HashSetData>();
            list.Add(new HashSetData() { Percentage = 35, Skill = "Java" });
            list.Add(new HashSetData() { Percentage = 40, Skill = ".Net" });
            list.Add(new HashSetData() { Percentage = 25, Skill = "Angular" });
            return list;
        }

        private async Task<string> SaveFile(IFormFile file)
        {
            var subDirectory = "upload";
            var target = Path.Combine(hostingEnvironment.ContentRootPath, subDirectory);

            if (!Directory.Exists(target))
            {
                Directory.CreateDirectory(target);
            }
            var files = Directory.GetFiles(target);
            foreach(var _f in files)
            {
               System.IO.File.Delete(_f);
            }
            if (file.Length <= 0) return string.Empty;
            var filePath = Path.Combine(target, file.FileName);
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }
            return filePath;
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
