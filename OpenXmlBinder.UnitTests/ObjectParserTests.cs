using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OpenXmlBinder.UnitTests
{
    public class ObjectParserTests
    {

        [Fact]
        public void GivenNonNestedObject_ObjectParser_FillsValidDictionnary()
        {
            dynamic obj = new
            {
                Test1 = "Hello",
                Test2 = "world !",
                Test_3 = 42,
                Test4_4 = 1.1,
                Test5=new DateTime(2022,07,03)
            };

            

            var parser = new ObjectParser();
            parser.ParseObject(obj);

            Assert.NotEmpty(parser.parseResult);
            Assert.Equal("Hello", parser.parseResult["Test1"]);
            Assert.Equal("world !", parser.parseResult["Test2"]);
            Assert.Equal("42", parser.parseResult["Test_3"]);
            Assert.Equal((1.1).ToString(), parser.parseResult["Test4_4"]);
            Assert.Equal(new DateTime(2022, 07, 03).ToString("D"), parser.parseResult["Test5"]);
        }

        [Fact]
        public void GivenNestedObject_ObjectParser_FillsValidDictionnary()
        {
            dynamic car = new
            {
                Brand = "Peugeot",
                Model = "206",
                Wheels = new
                {
                    Brand=new
                    {
                        Name="Michellin",
                        Origin="France"
                    },
                    Size=15
                },
            };

            var parser = new ObjectParser();
            parser.ParseObject(car);


            Assert.NotEmpty(parser.parseResult);
            Assert.Equal("Peugeot", parser.parseResult["Brand"]);
            Assert.Equal("206", parser.parseResult["Model"]);
            Assert.Equal("Michellin", parser.parseResult["Wheels.Brand.Name"]);
            Assert.Equal("France", parser.parseResult["Wheels.Brand.Origin"]);
            Assert.Equal("15", parser.parseResult["Wheels.Size"]);
        }
    }
}
