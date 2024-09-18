using System.Reflection.Emit;
using System.Reflection;

namespace Aplikace.Tridy
{
    //mělo by dynamicky vgenerovat třídu

    public class Generator
    {

        static void Main(string[] args)
        {
            // Seznam stringů obsahující názvy vlastností
            List<string> propertyNames = new List<string> { "Name", "Age", "Occupation" };

            // Vytvoření dynamické třídy
            Type dynamicClass = CreateDynamicClass(propertyNames);

            // Vytvoření instance dynamické třídy
            object classInstance = Activator.CreateInstance(dynamicClass);

            // Nastavení hodnot vlastností
            SetPropertyValue(classInstance, "Name", "Martin");
            SetPropertyValue(classInstance, "Age", 30);
            SetPropertyValue(classInstance, "Occupation", "Programmer");

            // Získání hodnot vlastností
            Console.WriteLine("Name: " + GetPropertyValue(classInstance, "Name"));
            Console.WriteLine("Age: " + GetPropertyValue(classInstance, "Age"));
            Console.WriteLine("Occupation: " + GetPropertyValue(classInstance, "Occupation"));
        }

        // Metoda pro dynamické vytvoření třídy
        public static Type CreateDynamicClass(List<string> propertyNames)
        {
            // Dynamický generátor třídy
            AssemblyName assemblyName = new AssemblyName("DynamicAssembly");
            AssemblyBuilder assemblyBuilder = AssemblyBuilder.DefineDynamicAssembly(assemblyName, AssemblyBuilderAccess.Run);
            ModuleBuilder moduleBuilder = assemblyBuilder.DefineDynamicModule("MainModule");

            // Vytvoření dynamické třídy
            TypeBuilder typeBuilder = moduleBuilder.DefineType("DynamicClass", TypeAttributes.Public);

            // Přidání vlastností do dynamické třídy
            foreach (var propertyName in propertyNames)
            {
                CreateProperty(typeBuilder, propertyName, typeof(object));
            }

            // Vytvoření typu (třídy)
            return typeBuilder.CreateType();
        }

        // Metoda pro vytvoření vlastnosti
        public static void CreateProperty(TypeBuilder typeBuilder, string propertyName, Type propertyType)
        {
            // Pole pro uložení hodnoty vlastnosti
            FieldBuilder fieldBuilder = typeBuilder.DefineField("_" + propertyName, propertyType, FieldAttributes.Private);

            // Vlastnost s getterem a setterem
            PropertyBuilder propertyBuilder = typeBuilder.DefineProperty(propertyName, PropertyAttributes.HasDefault, propertyType, null);

            // Vytvoření get metody
            MethodBuilder getMethodBuilder = typeBuilder.DefineMethod("get_" + propertyName, MethodAttributes.Public, propertyType, Type.EmptyTypes);
            ILGenerator getIL = getMethodBuilder.GetILGenerator();
            getIL.Emit(OpCodes.Ldarg_0);
            getIL.Emit(OpCodes.Ldfld, fieldBuilder);
            getIL.Emit(OpCodes.Ret);

            // Vytvoření set metody
            MethodBuilder setMethodBuilder = typeBuilder.DefineMethod("set_" + propertyName, MethodAttributes.Public, null, new Type[] { propertyType });
            ILGenerator setIL = setMethodBuilder.GetILGenerator();
            setIL.Emit(OpCodes.Ldarg_0);
            setIL.Emit(OpCodes.Ldarg_1);
            setIL.Emit(OpCodes.Stfld, fieldBuilder);
            setIL.Emit(OpCodes.Ret);

            // Přidání getteru a setteru k vlastnosti
            propertyBuilder.SetGetMethod(getMethodBuilder);
            propertyBuilder.SetSetMethod(setMethodBuilder);
        }

        // Nastavení hodnoty vlastnosti
        public static void SetPropertyValue(object obj, string propertyName, object value)
        {
            PropertyInfo prop = obj.GetType().GetProperty(propertyName);
            if (prop != null && prop.CanWrite)
            {
                prop.SetValue(obj, value, null);
            }
        }

        // Získání hodnoty vlastnosti
        public static object GetPropertyValue(object obj, string propertyName)
        {
            PropertyInfo prop = obj.GetType().GetProperty(propertyName);
            if (prop != null && prop.CanRead)
            {
                return prop.GetValue(obj, null);
            }
            return null;
        }
    }

}
