using System.Web.Mvc;
using System.Web.Security;

namespace iembee.Controllers
{
    public class AccountController : Controller
    {
        private readonly string UserName = "admin";
        private readonly string Password = "@admin";
        // GET: Account
        public ActionResult Index()
        {
            if (!string.IsNullOrEmpty(HttpContext.User.Identity.Name))
            {
                return RedirectToAction("Index", "Home");
            }
            return View();
        }
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Login(string username, string password, bool memberme = false)
        {
            if(username == UserName && password == Password)
            {
                FormsAuthentication.SetAuthCookie(username, memberme);
                return RedirectToAction("Index", "Home");
            }
            return View("Index");
        }
        [Authorize]
        public ActionResult Logout()
        {
            FormsAuthentication.SignOut();
            return View("Index");
        }
    }
}