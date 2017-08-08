using System.Collections.Generic;

namespace EdiConnectorService_C_Sharp
{
    public class Agent
    {
        private Queue<ICommand> commandsQueue = new Queue<ICommand>();

        public Agent()
        {
        }

        public void QueueCommand(ICommand _command)
        {
            commandsQueue.Enqueue(_command);
            _command.execute();
            commandsQueue.Dequeue();
        }    
    }
}
