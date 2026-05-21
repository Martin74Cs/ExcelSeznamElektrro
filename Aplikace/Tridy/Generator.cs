using System.Reflection;
using System.Reflection.Emit;

namespace Aplikace.Tridy
{
    //Příklad použití key/value úložiště:
    //Generator.Instance.Set("Projekt", "Rozvaděč A");
    //string? projekt = Generator.Instance.Get<string>("Projekt");
    //Console.WriteLine(projekt);

    //Použití runtime class:
    //Generator.Hlavni();

    public class Generator
    {
        // =========================
        // Singleton
        // =========================

        private static readonly Generator _instance = new();

        public static Generator Instance => _instance;

        private Generator()
        {
        }

        // =========================
        // Interní úložiště
        // =========================

        private readonly Dictionary<string, object> Seznam = [];

        // =========================
        // TEST
        // =========================

        public static void Hlavni()
        {
            List<string> propertyNames =
            [
                "Name",
                "Age",
                "Occupation"
            ];

            // vytvoření runtime typu
            Type dynamicClass =
                CreateDynamicClass(propertyNames);

            // instance
            object? classInstance =
                Activator.CreateInstance(dynamicClass);

            if (classInstance == null)
                return;

            // nastavení property
            SetPropertyValue(classInstance, "Name", "Martin");
            SetPropertyValue(classInstance, "Age", 30);
            SetPropertyValue(classInstance, "Occupation", "Programmer");

            // čtení property
            Console.WriteLine(
                "Name: " +
                GetPropertyValue(classInstance, "Name"));

            Console.WriteLine(
                "Age: " +
                GetPropertyValue(classInstance, "Age"));

            Console.WriteLine(
                "Occupation: " +
                GetPropertyValue(classInstance, "Occupation"));
        }

        // =========================
        // Dynamická třída
        // =========================

        public static Type CreateDynamicClass(
            List<string> propertyNames)
        {
            AssemblyName assemblyName =
                new("DynamicAssembly");

            AssemblyBuilder assemblyBuilder =
                AssemblyBuilder.DefineDynamicAssembly(
                    assemblyName,
                    AssemblyBuilderAccess.Run);

            ModuleBuilder moduleBuilder =
                assemblyBuilder.DefineDynamicModule(
                    "MainModule");

            TypeBuilder typeBuilder =
                moduleBuilder.DefineType(
                    "DynamicClass",
                    TypeAttributes.Public);

            foreach (var propertyName in propertyNames)
            {
                CreateProperty(
                    typeBuilder,
                    propertyName,
                    typeof(object));
            }

            return typeBuilder.CreateType()!;
        }

        // =========================
        // Vytvoření property
        // =========================

        public static void CreateProperty(
            TypeBuilder typeBuilder,
            string propertyName,
            Type propertyType)
        {
            // private field
            FieldBuilder fieldBuilder =
                typeBuilder.DefineField(
                    "_" + propertyName,
                    propertyType,
                    FieldAttributes.Private);

            // property
            PropertyBuilder propertyBuilder =
                typeBuilder.DefineProperty(
                    propertyName,
                    PropertyAttributes.HasDefault,
                    propertyType,
                    null);

            // ================= GET =================

            MethodBuilder getMethodBuilder =
                typeBuilder.DefineMethod(
                    "get_" + propertyName,
                    MethodAttributes.Public |
                    MethodAttributes.SpecialName |
                    MethodAttributes.HideBySig,
                    propertyType,
                    Type.EmptyTypes);

            ILGenerator getIL =
                getMethodBuilder.GetILGenerator();

            getIL.Emit(OpCodes.Ldarg_0);
            getIL.Emit(OpCodes.Ldfld, fieldBuilder);
            getIL.Emit(OpCodes.Ret);

            // ================= SET =================

            MethodBuilder setMethodBuilder =
                typeBuilder.DefineMethod(
                    "set_" + propertyName,
                    MethodAttributes.Public |
                    MethodAttributes.SpecialName |
                    MethodAttributes.HideBySig,
                    null,
                    [propertyType]);

            ILGenerator setIL =
                setMethodBuilder.GetILGenerator();

            setIL.Emit(OpCodes.Ldarg_0);
            setIL.Emit(OpCodes.Ldarg_1);
            setIL.Emit(OpCodes.Stfld, fieldBuilder);
            setIL.Emit(OpCodes.Ret);

            // přiřazení getter/setter
            propertyBuilder.SetGetMethod(getMethodBuilder);
            propertyBuilder.SetSetMethod(setMethodBuilder);
        }

        // =========================
        // Nastavení property
        // =========================

        public static void SetPropertyValue(
            object obj,
            string propertyName,
            object value)
        {
            PropertyInfo? prop =
                obj.GetType().GetProperty(propertyName);

            if (prop != null && prop.CanWrite)
            {
                prop.SetValue(obj, value);
            }
        }

        // =========================
        // Získání property
        // =========================

        public static object? GetPropertyValue(
            object obj,
            string propertyName)
        {
            PropertyInfo? prop =
                obj.GetType().GetProperty(propertyName);

            if (prop != null && prop.CanRead)
            {
                return prop.GetValue(obj);
            }

            return null;
        }

        // =========================
        // Key/Value úložiště
        // =========================

        public void Set(string klic, object hodnota)
        {
            Seznam[klic] = hodnota;
        }

        public T? Get<T>(string klic)
        {
            if (Seznam.TryGetValue(klic, out object? hodnota))
            {
                if (hodnota is T value)
                {
                    return value;
                }
            }

            return default;
        }
    }
}