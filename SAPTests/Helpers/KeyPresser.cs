using WindowsInput;
using WindowsInput.Native;

namespace SAPTests.Helpers
{
    public static class KeyPresser
    {
        private static InputSimulator inputSimulator = new InputSimulator();
        public static void PressKey(string key, int times = 1, int timeoutMs = 0)
        {
            VirtualKeyCode keyCode;

            // Map numeric strings and special keys to VirtualKeyCode values
            Dictionary<string, VirtualKeyCode> keyMap = new Dictionary<string, VirtualKeyCode>
        {
            { "0", VirtualKeyCode.VK_0 },
            { "1", VirtualKeyCode.VK_1 },
            { "2", VirtualKeyCode.VK_2 },
            { "3", VirtualKeyCode.VK_3 },
            { "4", VirtualKeyCode.VK_4 },
            { "5", VirtualKeyCode.VK_5 },
            { "6", VirtualKeyCode.VK_6 },
            { "7", VirtualKeyCode.VK_7 },
            { "8", VirtualKeyCode.VK_8 },
            { "9", VirtualKeyCode.VK_9 },
            { "PageDown", VirtualKeyCode.NEXT }
        };

            if (keyMap.TryGetValue(key, out keyCode) || Enum.TryParse(key, out keyCode))
            {
                InputSimulator inputSimulator = new InputSimulator();
                for (int i = 1; i <= times; i++)
                {
                    inputSimulator.Keyboard.KeyPress(keyCode);

                    if (timeoutMs > 0)
                    {
                        Thread.Sleep(timeoutMs);
                    }
                }
            }
            else
            {
                Console.WriteLine("Invalid key definition.");
            }
        }

        public static void PressKeys(params string[] keys)
        {
            var keyCodes = ParseKeyCodes(keys);
            if (keyCodes == null || keyCodes.Count == 0)
            {
                Console.WriteLine("Invalid key definitions.");
                return;
            }

            if (keyCodes.Count > 1)
            {
                var modifiers = keyCodes.Take(keyCodes.Count - 1).ToArray();
                var mainKey = keyCodes.Last();

                // Press and hold modifier keys
                foreach (var modifier in modifiers)
                {
                    inputSimulator.Keyboard.KeyDown(modifier);
                }

                // Press and release the main key
                inputSimulator.Keyboard.KeyPress(mainKey);

                // Release the modifier keys
                foreach (var modifier in modifiers.Reverse())
                {
                    inputSimulator.Keyboard.KeyUp(modifier);
                }
            }
            else
            {
                // Press and release the single key
                inputSimulator.Keyboard.KeyPress(keyCodes.First());
            }
        }

        public static void PressAndHoldKeys(params string[] keys)
        {
            var keyCodes = ParseKeyCodes(keys);
            if (keyCodes == null || keyCodes.Count == 0)
            {
                Console.WriteLine("Invalid key definitions.");
                return;
            }

            // Press and hold all keys
            foreach (var keyCode in keyCodes)
            {
                inputSimulator.Keyboard.KeyDown(keyCode);
            }
        }

        public static void ReleaseKeys(params string[] keys)
        {
            var keyCodes = ParseKeyCodes(keys);
            if (keyCodes == null || keyCodes.Count == 0)
            {
                Console.WriteLine("Invalid key definitions.");
                return;
            }

            // Release all keys
            foreach (var keyCode in keyCodes)
            {
                inputSimulator.Keyboard.KeyUp(keyCode);
            }
        }

        private static List<VirtualKeyCode> ParseKeyCodes(params string[] keys)
        {
            var keyCodes = new List<VirtualKeyCode>();
            foreach (var key in keys)
            {
                if (Enum.TryParse(key, true, out VirtualKeyCode keyCode))
                {
                    keyCodes.Add(keyCode);
                }
                else
                {
                    Console.WriteLine($"Invalid key definition: {key}");
                    return null;
                }
            }
            return keyCodes;
        }
    }
}
