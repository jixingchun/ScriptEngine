using System;
using System.Collections.Generic;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ScriptEngine.Tests
{
    [TestClass]
    public class EngineHostCompileTest : TestBase
    {
        [TestMethod]
        public void TypeUtils_DispatchInvokeMember_SetProperty()
        {
            User user = new User { Name = "张三" };
            object obj = TypeUtils.DispatchInvokeMember("Name",BindingFlags.SetProperty,null,user,new string[]{"李四"});

            Assert.AreEqual(user.Name, "李四");
        }

        #region 编译脚本测试

        [TestMethod]
        public void CompileScript_VB_Script_Only()
        {
            string script = @"
dim a
dim b
dim c
a=1
b=2
c=a+b

";
            using (EngineHost engineHost = new EngineHost())
            {
                IList<ScriptException> actual = engineHost.CompileScript(script);
                Assert.AreEqual(actual.Count, 0);
            }
        }

        [TestMethod]
        public void Compile_Script_Have_CSharp_Ojbect_Expect_No_Object_Error()
        {
            string script = @"
    Dim a(2)
    a(0) = ""lisi""
    User.SayHello(a(0))

    a(1) = 100
    User.Age = a(1)
    
    Test = User.ToString()

";
            using (EngineHost engineHost = new EngineHost())
            {
                IList<ScriptException> actual = engineHost.CompileScript(script);

                // 缺少对象: 'User'
                Assert.AreEqual(actual.Count, 1);
                Assert.IsTrue(actual[0].Description.Contains("User"));
            }
        }

        [TestMethod]
        public void Compile_Script_Have_CSharp_Ojbect_OK()
        {
            string script = @"
    Dim a(2)
    a(0) = ""lisi""
    User.SayHello(a(0))

    a(1) = 100
    User.Age = a(1)
    
    Test = User.ToString()
";
            using (EngineHost engineHost = new EngineHost())
            {
                string name = "User";
                User value = new User() { Name = "zhangsan", Age = 25 };
                engineHost.AddNamedItem(name, value);

                IList<ScriptException> actual = engineHost.CompileScript(script);
                Assert.AreEqual(actual.Count, 0);
            }
        }
        #endregion

        #region 编译表达式测试
        [TestMethod]
        [Description("编译表达式")]
        public void CompileExpression_Result_Is_Ok()
        {
            using (var engine = new EngineHost())
            {
                string expression = @"user.Age = user.Age  +  10";

                User user = new User { Name = "张三", Age = 25 };
                engine.AddNamedItem("user", user);

                IList<ScriptException> actual = engine.CompileScript(expression, true);
                Assert.AreEqual(0, actual.Count);
            }
        }

        [TestMethod]
        [Description("编译表达式")]
        public void CompileExpression_Result_Is_Ng()
        {
            using (var engine = new EngineHost())
            {
                string expression = @"user.Age  +  10";
                IList<ScriptException> actual = engine.CompileScript(expression, true);
                Assert.AreEqual(1, actual.Count);
                Assert.AreEqual("800A01A8", actual[0].Number.ToString("X8"));//缺少对象 user

            }
        }
        #endregion

        #region 执行表达式测试
        [TestMethod]
        [Description("编译表达式")]
        public void EvalExpression_Result_Is_Int()
        {
            using (var engine = new EngineHost())
            {
                User user = new User { Name = "Zhangsan", Age = 25 };
                engine.AddNamedItem("User", user);

                string expression = @"User.Age + 10";
                int expected = 35;
                object actual = engine.EvalExpression(expression);
                Assert.AreEqual(expected, Convert.ToInt32(actual));
            }
        }

        [TestMethod]
        [Description("编译表达式")]
        public void EvalExpression_Result_Is_Array()
        {
            using (var engine = new EngineHost())
            {
                User user = new User { Name = "Zhangsan", Age = 25 };
                engine.AddNamedItem("User", user);

                string expression = @"User.GetLanguageSkill()";
                object actual = engine.EvalExpression(expression);

                Assert.IsInstanceOfType(actual, (new string[] { }).GetType());

                string[] lan = (string[])actual;
                Assert.AreEqual(2, lan.Length);
                Assert.AreEqual("zh", lan[0]);
            }
        }

        //[TestMethod]
        [Description("编译表达式")]
        //[ExpectedException(typeof(Exception),"")]
        public void EvalExpression_Target_Script_Is_Not_Expression()
        {
            using (var engine = new EngineHost())
            {
                User user = new User { Name = "Zhangsan", Age = 25 };
                engine.AddNamedItem("User", user);

                string expression = @"
Sub Test()
  User.Age = User.Age
END Sub";

                object actual = engine.EvalMethod(expression, "Test", null);
                Assert.AreEqual(null, actual);
            }
        }
        #endregion

        #region 执行脚本函数测试

        [TestMethod]
        public void Excute_Function_No_Params()
        {
            string script = @"
Function Test()
    Dim a(2)
    a(0) = ""lisi""
    User.SayHello(a(0))

    a(1) = 100

    User.Name = a(0)
    User.Age = a(1)
    
    Test = User.ToString()
End Function
";
            using (EngineHost engineHost = new EngineHost())
            {
                string name = "User";
                User userInfo = new User() { Name = "zhangsan", Age = 25 };
                engineHost.AddNamedItem(name, userInfo);

                object actual = engineHost.EvalMethod(script, "Test");

                Assert.AreEqual(100, userInfo.Age);
                string expected = $"Name={"lisi"}\tAge={100}";
                Assert.AreEqual(expected, actual);
            }
        }

        [TestMethod]
        public void Excute_Function_Have_String_Params()
        {
            string script = @"
Function Test(userName)
    
    user.Name = userName
    Test = user.ToString()
    
End Function
";
            using (EngineHost engineHost = new EngineHost())
            {
                User userInfo = new User() { Name = "zhangsan", Age = 25 };
                engineHost.AddNamedItem("user", userInfo);

                string userName = "Lisi";
                object actual = engineHost.EvalMethod(script, "Test", userName);

                Assert.AreEqual("Lisi", userInfo.Name);

                string expected = $"Name={"Lisi"}\tAge={25}";
                Assert.AreEqual(expected, actual);
            }
        }

        [TestMethod]
        public void Excute_Function_Have_Class_Params()
        {
            string script = @"
Function Test(user)
    
    user.Name = ""Lisi""
    Test = user.ToString()
End Function
";
            using (EngineHost engineHost = new EngineHost())
            {
                User userInfo = new User() { Name = "zhangsan", Age = 25 };
                DispatchProxy dispatchProxy = engineHost.CreateDispatchProxy(userInfo);

                string userName = "Lisi";

                object actual = engineHost.EvalMethod(script, "Test", dispatchProxy);
                Assert.AreEqual(userName, userInfo.Name);
            }
        }

        [TestMethod]
        public void Excute_Function_Have_String_And_Class_Params()
        {
            string script = @"
Function Test(userName,user)
    
    user.Name = userName
    user.Age = 10
    user.Age = user.Age + 10
    Test = user.ToString()
End Function
";
            using (EngineHost engineHost = new EngineHost())
            {
                string userName = "Lisi";
                User userInfo = new User() { Name = "zhangsan", Age = 25 };
                DispatchProxy dispatchProxy = engineHost.CreateDispatchProxy(userInfo);

                object actual = engineHost.EvalMethod(script, "Test", userName, dispatchProxy);

                Assert.AreEqual("Lisi", userInfo.Name);
                Assert.AreEqual(20, userInfo.Age);
            }
        }

        [TestMethod]
        public void Excute_Function_Call_Other_Script()
        {
            string script = @"
Function Test()
    dim a(1)
    a(0) = ""football""
    a(1) = ""baskedball""
    Test = ModuleA.Run((a))
End Function
";
            using (EngineHost engineHost = new EngineHost())
            {
                ModuleA moduleA = new ModuleA();
                engineHost.AddNamedItem("ModuleA", moduleA);
                object actual = engineHost.EvalMethod(script, "Test");
                Assert.AreEqual("footballbaskedball", actual);
            }
        }

        //[TestMethod]
        public void Excute_Function_Simple_Array_Input()
        {
            string script = @"
        Function Test(names)
            dim msg
            If IsArray( names ) Then
            For i = 0 To UBound(names) Step 1
                msg = msg & names(i)
            Next
            Else
                msg = ""is not array ""
            End if
            Test = msg
        End Function
        ";
            using (EngineHost engineHost = new EngineHost())
            {
                object balls = new string[2] { "football", "basketball" };

                List<string> list = new List<string>() { "football", "basketball" };

                //DispatchProxy dispatchProxy = engineHost.CreateDispatchProxy(list);

                DispatchProxy dispatchProxy = engineHost.CreateDispatchProxy(balls);

                object actual = engineHost.EvalMethod(script, "Test", dispatchProxy);
                Assert.AreEqual("footballbasketball", balls);
            }
        }

        [TestMethod]
        public void Excute_Function_Have_String_And_Class_Params_xxx()
        {
            string script = @"
Function Test(userName,user)
    
    user.Name = userName
    user.Age = 10
    user.Age = user.Age + 10
    Test = user.Age
End Function
";
            using (EngineHost engineHost = new EngineHost())
            {
                string userName = "Lisi";
                User userInfo = new User() { Name = "zhangsan", Age = 25 };
                DispatchProxy dispatchProxy = engineHost.CreateDispatchProxy(userInfo);

                object actual = engineHost.EvalMethod(script, "Test", userName, dispatchProxy);

                Assert.AreEqual(20, actual);
            }
        }


        [TestMethod]
        public void Excute_Sub_Have_String_And_Class_Params()
        {
            string script = @"
Sub Test(userName,user)
    
    user.Name = userName
    user.Age = 10
    user.Age = user.Age + 10

End Sub
";
            using (EngineHost engineHost = new EngineHost())
            {
                string userName = "Lisi";
                User userInfo = new User() { Name = "zhangsan", Age = 25 };
                DispatchProxy dispatchProxy = engineHost.CreateDispatchProxy(userInfo);

                object actual = engineHost.EvalMethod(script, "Test", userName, dispatchProxy);

                Assert.AreEqual(null, actual);
                Assert.AreEqual(20, userInfo.Age);
            }
        }

        #endregion

    }
}
