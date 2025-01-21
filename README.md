# TaskMng

TaskMng is a task management application written in C#. This project includes various functionalities such as logging, configuration management, encryption, and more.

## Project Structure
TaskMng.sln 
taskrun/ 
BackupProgram.cs 
BackupTask.cs 
BackupUser.cs 
ConfigMng.cs 
CopyMng.cs 
EncryptMng.cs 
LogMng.cs 
taskrun.csproj

## Getting Started

### Prerequisites

- .NET 8.0 SDK or later
- Visual Studio or any other C# IDE

### Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/yourusername/TaskMng.git
    cd TaskMng/taskrun
    ```

2. Restore the dependencies:
    ```sh
    dotnet restore
    ```

### Building the Project

To build the project, run the following command in the [taskrun](http://_vscodecontentref_/7) directory:

```sh
dotnet build
---

### Running the Project
To run the project, use the following command:
---sh
dotnet run
---

## Usage
### Logging
The logging functionality is managed by the LogMng class. To initialize the logger, call the Initialize method with the desired log directory:

### Configuration Management
The configuration management is handled by the ConfigMng class. To log a message, use the Log method:

### Contributing
Contributions are welcome! Please fork the repository and create a pull request with your changes.

## License
This project is licensed under the MIT License - see the LICENSE file for details.

