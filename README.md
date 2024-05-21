# C# .net-Program-to-fetch-data-from-PostgreSQL
Source code for a desktop app which fetches data from PostgreSQL. Feel free to use refer. Note: comments wont be available for everything.

**-> Note: This program is only used for fetching data and exporting it in the format xls and csv, it cannot be used for writing data**

-> Special Feature: This program can automatically detect the list of database existing in your server and the list of tables present inside a database. **Note: Only public tables are visible not private or the built in tables**

This C# .NET program designed to fetch data from a PostgreSQL database typically involves establishing a connection to the database using a connection string, executing SQL commands, and handling the data retrieved. 

The program would utilize the Npgsql library, which is an open-source .NET Data Provider for PostgreSQL. It's fully integrated with the ADO.NET framework, which means it provides a seamless experience for interacting with the database. 

The process begins with defining a connection string that includes the server location, port, database name, user ID, and password. Once the connection is established, SQL commands are executed using the `NpgsqlCommand` class. Data retrieval can be performed using the `NpgsqlDataReader` or by filling a `DataSet` with the `NpgsqlDataAdapter`. 

The program can handle various operations such as querying data with SELECT statements, inserting new records, updating existing ones, and deleting records as needed. 

Error handling is crucial to manage any exceptions that may occur during database operations. Additionally, the program would include mechanisms for parameterized queries to prevent SQL injection attacks, ensuring the application's security.

Overall, the program serves as a robust and secure way to interact with a PostgreSQL database, providing the necessary functionality to perform operations efficiently. 
