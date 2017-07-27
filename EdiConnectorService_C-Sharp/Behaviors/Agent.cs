using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

namespace EdiConnectorService_C_Sharp
{
    public class Agent
    {
        private Queue<Command> commandsQueue = new Queue<Command>();

        public Agent()
        {
        }

        public void QueueCommand(Command _command)
        {
            commandsQueue.Enqueue(_command);
            _command.execute();
            commandsQueue.Dequeue();
        }    
    }
}
