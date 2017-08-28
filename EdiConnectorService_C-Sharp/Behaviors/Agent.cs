using System.Collections.Generic;

namespace EdiConnectorService_C_Sharp
{
    /// <summary>
    /// An agent object to issue commands
    /// </summary>
    public class Agent
    {
        // Queue object which holds all the queued up commands that are issued
        private Queue<ICommand> commandsQueue = new Queue<ICommand>();

        public Agent()
        {
        }

        /// <summary>
        /// Adds a command to the queue.
        /// </summary>
        /// <param name="_command">The command.</param>
        public void QueueCommand(ICommand _command)
        {
            commandsQueue.Enqueue(_command);
            _command.execute();
            commandsQueue.Dequeue();
        }    
    }
}
