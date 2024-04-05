#r "nuget: Selenium.WebDriver, 4.16.2"
#r "nuget: Selenium.Support, 4.16.1"
#r "nuget: DotNetSeleniumExtras.WaitHelpers, 3.11.0"
#r "nuget: NAudio, 2.2.1"


using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.Threading;
using SeleniumExtras.WaitHelpers;
using OpenQA.Selenium.Support.UI;
using NAudio.Wave;
using System.Text.RegularExpressions;
using OpenQA.Selenium.Interactions;

public string storeCode = "f9k7m";
public int x1 = 3;
public int x2 = 4;

public static Dictionary<string, string> LettersDictionary = new Dictionary<string, string>
{
    ["۰"] = "0",
    ["۱"] = "1",
    ["۲"] = "2",
    ["۳"] = "3",
    ["۴"] = "4",
    ["۵"] = "5",
    ["۶"] = "6",
    ["۷"] = "7",
    ["۸"] = "8",
    ["۹"] = "9"
};


await RunCode();

public async Task RunCode()
{
    
    // Specify the path to the ChromeDriver executable
    string chromeDriverPath = "chromedriver.exe"; // Update this with your actual path
    var chromeOptions = new ChromeOptions();
    string userName = System.Security.Principal.WindowsIdentity.GetCurrent().Name.Split("\\")[1];
    chromeOptions.AddArguments(@"user-data-dir=C:\Users\" + userName + @"\AppData\Local\Google\Chrome\User Data\" + storeCode);

    // Create an instance of ChromeDriver
    IWebDriver driver = new ChromeDriver(chromeDriverPath, chromeOptions);




    try
    {

        // Step 1
        driver.Navigate().GoToUrl("https://seller.digikala.com/pwa/consignment?page=1&size=200");

        await Task.Delay(5000);


        IWebElement checkbox = driver.FindElement(By.CssSelector("[class*='Checkbox__checkbox']"));
        checkbox.Click();

        IWebElement step1CreateButton = driver.FindElement(By.CssSelector("[class*='TableHeadAction__createPackageAction']"));
        step1CreateButton.Click();

        //Step 2
        await Task.Delay(5000);
        IReadOnlyCollection<IWebElement> rows = driver.FindElements(By.CssSelector("[class*='TableRowContainer__normal']"));

        // Output the found elements
        foreach (var row in rows)
        {
            Actions actions = new Actions(driver);
            // Get the Y-coordinate of the element
            int yOffset = row.Location.Y;
            // Execute JavaScript to scroll to the element's position with an additional offset
            ((IJavaScriptExecutor)driver).ExecuteScript("window.scrollTo(0, arguments[0] - 100);", yOffset); // Adjust the offset as needed
            await Task.Delay(1000);

            IReadOnlyCollection<IWebElement> divElements = row.FindElements(By.TagName("div"));
            int maxAllowedSend = PersianNumberToEnglish(divElements.ElementAt(12).Text);
            int totalOrders = PersianNumberToEnglish(divElements.ElementAt(10).Text);

            IWebElement availabilityElement = divElements.ElementAt(11);
            IReadOnlyCollection<IWebElement> pElements = availabilityElement.FindElements(By.TagName("p"));
            int pendingAvailability = PersianNumberToEnglish(pElements.ElementAt(0).Text.ExtractNumber());
            int stockAvailability = PersianNumberToEnglish(pElements.ElementAt(1).Text.ExtractNumber());
            int totalAvailability = pendingAvailability + stockAvailability;
            IWebElement sentCountInput = row.FindElement(By.CssSelector("[class*='IncrementInput__input-']"));
            //int sentCountValue = PersianNumberToEnglish(sentCountInput.GetAttribute("value"));
            int sendAmount = x2 * totalOrders - totalAvailability;

            if (sendAmount > maxAllowedSend) sendAmount = maxAllowedSend;
            if (sendAmount == maxAllowedSend) sendAmount = maxAllowedSend - 1;
            // Check if sendAmount is odd
            if (sendAmount % 2 != 0)
            {
                sendAmount--; // If odd, make it even by subtracting 1
            }
            Console.WriteLine($"{totalOrders} {pendingAvailability} {stockAvailability} {maxAllowedSend} {sendAmount}");

            if (maxAllowedSend < 4 || totalAvailability > x1 * totalOrders || sendAmount < 4)
            {
                IWebElement removeButtonElement = row.FindElement(By.TagName("button"));
                removeButtonElement.Click();
                await Task.Delay(1000);

                IReadOnlyCollection<IWebElement> popupButtons = driver.FindElements(By.CssSelector("[class*='uikit-collections-Popup-templates-confirm-confirmTemplate']"));
                popupButtons.ElementAt(1).Click();
                continue;
            }

            sentCountInput.Click();

            // Perform Ctrl+A (select all)
            actions.KeyDown(Keys.Control).SendKeys("a").KeyUp(Keys.Control).Perform();
            // Perform Backspace to delete the selected text
            actions.SendKeys(Keys.Backspace).Perform();

            sentCountInput.SendKeys(sendAmount.ToString());

            // Perform any actions with the elements as needed
        }

    }
    catch (Exception ex)
    {
        Console.WriteLine($"An error occurred: {ex.Message}");
    }
    finally
    {
        PlayNotificationSound();
        // Close the browser window
        //driver.Quit();
    }

}

public void PlayNotificationSound()
{
    string audioFilePath = "audiofile.wav";

    try
    {
        using (var audioFile = new AudioFileReader(audioFilePath))
        using (var outputDevice = new WaveOutEvent())
        {
            outputDevice.Init(audioFile);
            outputDevice.Play();
            Console.WriteLine("Playing audio. Press any key to exit...");
            Console.ReadKey();
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"An error occurred: {ex.Message}");
    }
}

public static int PersianNumberToEnglish(this string persianStr)
{

    return int.Parse(LettersDictionary.Aggregate(persianStr, (current, item) =>
                 current.Replace(item.Key, item.Value)));
}

public static string ExtractNumber(this string text)
{
    string numberPattern = @"\d+";
    Match match = Regex.Match(text, numberPattern);

    if (match.Success)
    {
        string numberString = match.Value;
        //int number = int.Parse(numberString);
        return numberString;
    }

    return "0"; // Return -1 if no number is found in the text
}