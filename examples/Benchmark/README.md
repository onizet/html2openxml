# Benchmarks

How to run the benchmark tool.
First build the project: `dotnet build -c Release`

Then run the performance test targeting multiple runtimes:
`dotnet run -c Release -f net8.0 --runtimes net48 net8.0`
