using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;

namespace ScriptEngine.Tests
{
    [TestClass]
    public class UtilsTests : TestBase
    {
        [TestMethod]
        public void TypeUtils_DispatchInvokeMember_SetProperty()
        {
            User user = new User { Name = "张三" };
            object obj = TypeUtils.DispatchInvokeMember("Name", BindingFlags.SetProperty, null, user, new string[] { "李四" });

            Assert.AreEqual(user.Name, "李四");
        }
    }
}
