using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScriptEngine.Tests
{
    public class User
    {
        public string Name { get; set; }
        private int _age;
        public int Age
        {
            get
            {
                Console.WriteLine($"Age Get = {_age}");
                return _age;
            }
            set
            {
                _age = value;
                Console.WriteLine($"Age Set = {_age}");
            }
        }


        public void SayHello(string message)
        {
            string ret = $"Helll{message}";
            Console.WriteLine(ret);
        }

        public override string ToString()
        {
            string ret = $"Name={Name}\tAge={Age}";
            Console.WriteLine(ret);
            return ret;
        }

        public void PlayBall()
        {
            string[] ball = new string[2] { "football", "basketball" };
            ModuleA moduleA = new ModuleA();
            moduleA.Run(ball);
        }

        public string[] GetLanguageSkill()
        {
            return new string[] { "zh", "en" };
        }
    }

    public class ModuleA
    {
        public string Run(string[] names)
        {
            using (EngineHost engine = new EngineHost())
            {
                string script = @"
            function ExcuteMethod(names)
                dim msg
            	If IsArray( names ) Then
            	For i = 0 To UBound(names) Step 1
            	    msg = msg & names(i)
            	Next
            	Else
            		msg = names
            	End if
                ExcuteMethod = msg
            end function";

                object obj = engine.EvalMethod(script, "ExcuteMethod", names);


                Console.Write($"ModuleA.ExcuteMethod result = {obj}");

                return obj.ToString();
            }
        }
    }
}
