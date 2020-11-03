ARG CLI_VERSION=latest
FROM microsoft/powershell:latest
RUN apt-get update && apt-get install -y curl sudo
RUN curl -sL https://deb.nodesource.com/setup_12.x | sudo -E bash -
RUN apt install nodejs -y
RUN npm i -g @pnp/cli-microsoft365@${CLI_VERSION} --production
RUN pwsh -c 'm365 cli completion pwsh setup --profile /root/.config/powershell/Microsoft.PowerShell_profile.ps1'
CMD [ "pwsh" ]