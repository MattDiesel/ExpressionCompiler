using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.CodeDom.Compiler;
using Microsoft.CSharp;
using System.CodeDom;
using System.Reflection;

namespace ExpressionCompiler
{
    /// <summary>
    /// A runtime compiled expression, taking a number of arguments.
    /// </summary>
    /// <typeparam name="T">The type for all arguments and return type. This is designed to be a
    /// numeric type, but would work with any type.</typeparam>
    public class Expression<T>
    {
        private static readonly string NamespaceName = "CustomExpression";
        private static readonly string ClassName = "Calculator";
        private static readonly string MethodName = "Calculate";


        private string expr;
        private string[] vars;
        private MethodInfo method;


        /// <summary>
        /// Creates a new expression with the string representation and variables
        /// </summary>
        /// <param name="expr">The string expression. This is actually a C# expression, rather than a
        /// maths statement</param>
        /// <param name="vars">An array of the expressions arg</param>
        public Expression(string expr, params string[] vars)
        {
            this.expr = expr;
            this.vars = vars;

            // Compile
            CompilerResults ret = this.Compile();

            // Check for errors during compiling
            if (ret.Errors.HasErrors)
            {
                throw new NotImplementedException("Expression<T>.Expression(): Error checking");
            }

            this.method = ret.CompiledAssembly
                .GetType(String.Join(".", Expression<T>.NamespaceName, Expression<T>.ClassName))
                .GetMethod(Expression<T>.MethodName);
        }


        /// <summary>
        /// Returns the expression as a string.
        /// </summary>
        /// <returns>The expression as a string.</returns>
        public override string ToString()
        {
            return this.expr;
        }


        /// <summary>
        /// Runs the expression, given the arguments, and returns the result.
        /// </summary>
        /// <param name="args">A dictionary of argments in the form {Parameter Name} => {Value}</param>
        /// <returns>The result of the expression.</returns>
        /// <remarks>This function is just a wrapper around <see cref="Invoked(params T[])"/>, so if
        /// the argument order is known that function is preferred.</remarks>
        public T Invoke(Dictionary<string, T> args)
        {
            T ret; 

            if (this.vars.Length == 0)
            {
                if (args.Count != 0)
                    throw new ArgumentException(String.Format("{0} additional parameters were supplied: {1}", args.Count, String.Join(",", args.Keys)));

                ret = this.Invoke();
            }
            else
            {
                T[] nargs = new T[this.vars.Length];

                for (int i = 0; i < this.vars.Length; i++)
                {
                    if (args.ContainsKey(this.vars[i]))
                    {
                        nargs[i] = args[this.vars[i]];
                        args.Remove(this.vars[i]);
                    }
                    else
                        throw new ArgumentException(String.Format("The parameter '{0}' was not supplied.", this.vars[i]));
                }

                if (args.Count != 0)
                    throw new ArgumentException(String.Format("{0} additional parameters were supplied: {1}", args.Count, String.Join(",", args.Keys)));

                ret = this.Invoke(nargs);
            }

            return ret;
        }


        /// <summary>
        /// Runs the expression, given the arguments, and returns the result.
        /// </summary>
        /// <param name="args">A list of argument values. These need to be in the same order as the variable
        /// names were supplied in the constructor.</param>
        /// <returns>The result of the expression.</returns>
        /// <remarks>If the argument order is not known then use the
        /// <see cref="Invoke(Dictionary<String,T>)"/> method instead.</remarks>
        public T Invoke(params T[] args)
        {
            if (args.Length != this.vars.Length)
                throw new ArgumentException(String.Format("Incorrect number of arguments given. Expected {0}, got {1}.", this.vars.Length, args.Length));

            object[] oargs = new object[args.Length];
            args.CopyTo(oargs, 0);

            return (T)this.method.Invoke(null, oargs);
        }

        /// <summary>
        /// Compiles the expression based on the expr and vars members
        /// </summary>
        /// <returns>The results of compiling. Error checking is the responibility of the invoker.</returns>
        private CompilerResults Compile()
        {
            // Build the DOM tree, namespace:
            CodeNamespace myNamespace = new CodeNamespace(Expression<T>.NamespaceName);
            myNamespace.Imports.Add(new CodeNamespaceImport("System"));

            // Build the class declaration	
            CodeTypeDeclaration classDeclaration = new CodeTypeDeclaration(Expression<T>.ClassName);
            classDeclaration.IsClass = true;
            classDeclaration.Attributes = MemberAttributes.Public | MemberAttributes.Static;
            myNamespace.Types.Add(classDeclaration);

            // Build the Method
            CodeMemberMethod myMethod = new CodeMemberMethod();
            myMethod.Name = Expression<T>.MethodName;
            myMethod.Attributes = MemberAttributes.Public | MemberAttributes.Static;
            myMethod.ReturnType = new CodeTypeReference(typeof(T));
            myMethod.Statements.Add(new CodeMethodReturnStatement(new CodeSnippetExpression(expr)));

            foreach (string s in this.vars)
                myMethod.Parameters.Add(new CodeParameterDeclarationExpression(typeof(T), s));

            classDeclaration.Members.Add(myMethod);

            // Finally the compile unit:
            CodeCompileUnit codeUnit = new CodeCompileUnit();
            codeUnit.Namespaces.Add(myNamespace);
            codeUnit.ReferencedAssemblies.Add("mscorlib.dll");
            codeUnit.ReferencedAssemblies.Add("System.dll");

            CodeDomProvider codeProvider = new CSharpCodeProvider();
            
            CompilerParameters parms = new CompilerParameters();
            parms.CompilerOptions = "/target:library /optimize";
            parms.GenerateExecutable = false;
            parms.GenerateInMemory = true;
            parms.IncludeDebugInformation = false;
            parms.ReferencedAssemblies.Add("mscorlib.dll");
            parms.ReferencedAssemblies.Add("System.dll");

            return codeProvider.CompileAssemblyFromDom(parms, codeUnit);
        }
    }
}
