namespace STCU.HEApplicationsExtract.Core.Services
{
    using System.Configuration;
    using Serilog;
    using Serilog.Core;
    using Serilog.Core.Enrichers;
    using Serilog.Enrichers;

    /// <summary>
    /// Logging implementation for the recording of application events.
    /// </summary>
    public static class LogConfiguration
    {
        #region Methods

        //Log.CloseAndFlush();
        /// <summary>
        /// Configure a generic logging instance to be used across the whole application.
        /// </summary>
        public static void ConfigureSerilog()
        {
            Log.Logger = new LoggerConfiguration()
                .ReadFrom.AppSettings()
                .Enrich.With(GetEnrichers())
                .CreateLogger();
        }

        public static void FlushSerilog()
        {
            Log.CloseAndFlush();
        }

        #endregion

        #region PrivateMethods

        private static ILogEventEnricher[] GetEnrichers()
        {
            var enrichers = new ILogEventEnricher[]
            {
                new MachineNameEnricher(),
                new PropertyEnricher("ApplicationName", ConfigurationManager.AppSettings["Application.Name"]),
                new PropertyEnricher("Environment", ConfigurationManager.AppSettings["Environment"])
            };

            return enrichers;
        }

        #endregion
    }
}
