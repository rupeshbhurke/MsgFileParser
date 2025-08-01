FROM mcr.microsoft.com/dotnet/runtime-deps:8.0
WORKDIR /app

# Install Linux dependencies for MSGReader
RUN apt-get update && apt-get install -y libgdiplus libc6-dev && rm -rf /var/lib/apt/lists/*

# Copy published files from local 'publish' directory
COPY publish/ .

ENTRYPOINT ["./MsgFileParser"]