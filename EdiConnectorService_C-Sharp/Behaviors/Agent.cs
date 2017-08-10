using System.Collections.Generic;
using System.Threading;

namespace EdiConnectorService_C_Sharp
{
    public class Agent
    {
        private Queue<ICommand> commandsQueue = new Queue<ICommand>();
        readonly object locker = new object();

        public Agent()
        {
        }

        public void QueueCommand(ICommand _command)
        {
            commandsQueue.Enqueue(_command);
        }    

        public void ExecuteCommand(ICommand _command)
        {
            _command.execute();
        }

        public void ExecuteCommandQueue()
        {
            foreach (ICommand command in commandsQueue)
            {
                ThreadPool.GetAvailableThreads(out var availableThreads, out var i);
                if (availableThreads > 0)
                    ThreadPool.QueueUserWorkItem(new WaitCallback(CommandWorkerThread), command);
            }
        }

        private void CommandWorkerThread(System.Object _command)
        {
            lock (locker)
            {
                ICommand command = (ICommand)_command;
                command.execute();
                commandsQueue.Dequeue();
            }
        }
    }
}
