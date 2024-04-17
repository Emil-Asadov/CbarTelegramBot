#See https://aka.ms/customizecontainer to learn how to customize your debug container and how Visual Studio uses this Dockerfile to build your images for faster debugging.

FROM mcr.microsoft.com/dotnet/runtime:6.0 AS base
WORKDIR /app

FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build
WORKDIR /src
COPY ["ConsoleAppCbarTelegramBot/ConsoleAppCbarTelegramBot.csproj", "ConsoleAppCbarTelegramBot/"]
RUN dotnet restore "ConsoleAppCbarTelegramBot/ConsoleAppCbarTelegramBot.csproj"
COPY . .
WORKDIR "/src/ConsoleAppCbarTelegramBot"
RUN dotnet build "ConsoleAppCbarTelegramBot.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "ConsoleAppCbarTelegramBot.csproj" -c Release -o /app/publish /p:UseAppHost=false

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "ConsoleAppCbarTelegramBot.dll"]