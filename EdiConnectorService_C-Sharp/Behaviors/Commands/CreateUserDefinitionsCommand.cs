namespace EdiConnectorService_C_Sharp
{
    /// <summary>
    /// Concrete Command class
    /// This command is used to create the user defined tables and fields for every connected server
    /// </summary>
    /// <seealso cref="EdiConnectorService_C_Sharp.Command" />
    public class CreateUserDefinitionsCommand : ICommand
    {
        public CreateUserDefinitionsCommand()
        {

        }

        public void execute()
        {
            // Iterate through every connected server
            foreach (string connectedServer in ConnectionManager.getInstance().GetAllConnectedServers())
            {
                // Create user defined tables for this connection
                UserDefined.CreateTables(connectedServer);
                // Create user defined fields for this connection
                UserDefined.CreateFields(connectedServer);
            }
        }
    }
}
