FROM mcr.microsoft.com/dotnet/sdk:9.0-alpine AS build
WORKDIR /build
COPY . /build
RUN mkdir /app

RUN dotnet restore /build/VisioStencilCreator.App/VisioStencilCreator.App.csproj
RUN dotnet publish /build/VisioStencilCreator.App/VisioStencilCreator.App.csproj -o /app

FROM mcr.microsoft.com/dotnet/aspnet:9.0-alpine AS final
COPY --from=build /app /app

ENTRYPOINT [ "dotnet", "/app/VisioStencilCreator.App.dll" ]
