//##############################  EXAMPLE CODE ##############################
//public static void Main(string[] args)
//{
//    Option[] options =
//    {
//                new Option("Option 1, Say Hello", SayHello),
//                new Option("Option 4, Quit")
//            };
//    SDMenu m = new SDMenu(options, "Example Menu");
//    m.Activate();
//}

//public static void SayHello()
//{
//    Console.WriteLine("Hello World!");
//    Console.ReadLine();
//}
//###########################################################################
namespace LibSrd_NETCore
{
    public class Option
    {
        public string Text;
        public static int position = -1;
        public Action Method;
        public ConsoleColor Color;

        /// <summary>
        /// A menu item for SDMenu.
        /// </summary>
        /// <param name="text">How you the option will appear on the menu.</param>
        /// <param name="method">The action you would like the option to perform upon click.</param>
        /// <param name="color">An optional parameter where you can change the colour of the text displayed on the screen</param>
        public Option(string text, Action method = null, ConsoleColor color = ConsoleColor.White)
        {
            Text = text;
            if (method != null)
                Method = method;
            position++;
            Color = color;
        }
    }

    public class SDMenu
    {
        //Public so the menu can be edited after decleration.
        public Option[] Options;
        public string MenuName;
        private int MenuLength;
        private string TopString;

        /// <summary>
        /// Creates a console based menu that can be acivated with menu.Activate();
        /// </summary>
        /// <param name="options">A list of Options</param>
        /// <param name="menuName">The name of the menu that will be displayed.</param>
        public SDMenu(Option[] options, string menuName = null)
        {
            MenuName = menuName;
            Options = options;
        }

        /// <summary>
        /// Will draw the menu to the screen and activate its functionality.
        /// </summary>
        public void Activate()
        {
            Console.Clear();
            int cursor = 0;
            MenuLength = menuLength();
            TopString = calculateTop();

            while (true)
            {
                Draw(cursor); //Draw the menu
                ConsoleKey keyPressed = Console.ReadKey().Key; //Record next keypress

                if (keyPressed == ConsoleKey.UpArrow && cursor > 0)
                    cursor--; //Up

                else if (keyPressed == ConsoleKey.DownArrow && cursor < Options.Length - 1)
                    cursor++; //Down

                else if (keyPressed == ConsoleKey.Enter) //Confirm
                {
                    if (Options[cursor].Method != null)
                    {
                        Console.Clear();
                        Options[cursor].Method();
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Draws the menu graphic to the console.
        /// </summary>
        /// <param name="cursor">The position through the list that is selected. (Zero Based)</param>
        private void Draw(int cursor)
        {
            //Clear the console and draw the top of the menu.
            Console.SetCursorPosition(0, 0);
            Console.WriteLine(TopString);

            if (MenuName != null) //Write the menu name
            {
                //Menu Name
                Console.Write($"|  {MenuName}");
                Console.SetCursorPosition(MenuLength, 1);
                Console.WriteLine(" |");

                //Gap Line
                Console.Write("|");
                Console.SetCursorPosition(MenuLength, 2);
                Console.WriteLine(" |");
            }

            for (int i = 0; i < Options.Length; i++) //Drawing each option
            {
                if (cursor == i) //If the current item is selected.
                {
                    Console.Write("|> ");
                    Console.ForegroundColor = ConsoleColor.Black;
                    Console.BackgroundColor = ConsoleColor.White;
                    Console.Write(Options[i].Text);
                    Console.BackgroundColor = ConsoleColor.Black;
                    Console.ForegroundColor = ConsoleColor.White;
                }
                else //Normal option
                {
                    Console.Write("|  ");
                    Console.ForegroundColor = Options[i].Color;
                    Console.Write(Options[i].Text);
                    Console.ForegroundColor = ConsoleColor.White;
                }
                Console.SetCursorPosition(MenuLength, i + 3);
                Console.WriteLine(" |");
            }
            Console.WriteLine(TopString); //Draw the bottom of the menu
        }

        /// <summary>
        /// finds the max length of the menu, using option lengths and title length.
        /// </summary>
        private int menuLength()
        {
            int length = 0;
            foreach (Option option in Options)
                if (option.Text.Length > length) length = option.Text.Length; //Checking max option

            if (MenuName != null && MenuName.Length > length) length = MenuName.Length; //Checking title

            return length + 4;
        }

        /// <summary>
        /// Calculating an appropriately sized Bar wherever called. I.e (+--------------+)
        /// </summary>
        private string calculateTop()
        {
            string top = "+";

            for (int i = 0; i < MenuLength; i++)
                top += "-";

            return top + "+";
        }
    }
}