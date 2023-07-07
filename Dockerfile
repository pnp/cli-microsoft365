FROM mcr.microsoft.com/powershell:alpine-3.17

ARG CLI_VERSION=latest

LABEL name="m365pnp/cli-microsoft365:${CLI_VERSION}" \
  description="Manage Microsoft 365 and SharePoint Framework projects on any platform" \
  homepage="https://pnp.github.io/cli-microsoft365" \
  maintainers="Waldek Mastykarz <waldek@mastykarz.nl>, \
  Garry Trinder <garry.trinder@live.com>, \
  Albert-Jan Schot <appie@digiwijs.nl>, \
  Rabia Williams <rabiawilliams@gmail.com>, \
  Patrick Lamber <patrick@nubo.eu>, \
  Arjun Menon <arjun.umenon@gmail.com>, \
  Adam Wojcik <adam.wojcik.it@gmail.com>, \
  Martin Lingstuyl <mlingstuyl@live.com>, \
  Jasey Waegebaert <jaseyw@gmigroup.be>, \
  Milan Holemans <Milan.Holemans@vanroey.be>" \
  com.azure.dev.pipelines.agent.handler.node.path="/usr/bin/node"

RUN apk add --no-cache \
  curl \
  sudo \
  bash \
  shadow \
  bash-completion \
  nodejs \
  npm \
  python3 \
  py3-pip

RUN adduser --system cli-microsoft365
USER cli-microsoft365

WORKDIR /home/cli-microsoft365

ENV 0="/bin/bash" \
  SHELL="bash" \
  NPM_CONFIG_PREFIX=/home/cli-microsoft365/.npm-global \
  PATH=$PATH:/home/cli-microsoft365/.npm-global/bin:/home/cli-microsoft365/.local/bin \
  CLIMICROSOFT365_ENV="docker"

RUN bash -c 'echo "export PATH=$PATH:/home/cli-microsoft365/.npm-global/bin:/home/.local/bin" >> ~/.bash_profile' \
  && bash -c 'echo "export CLIMICROSOFT365_ENV=\"docker\"" >> ~/.bash_profile' \
  && bash -c 'npm i -g @pnp/cli-microsoft365@${CLI_VERSION} --production --quiet --no-progress' \ 
  && bash -c 'echo "source /etc/profile.d/bash_completion.sh" >> ~/.bash_profile' \
  && bash -c 'echo "alias \"m365?\"=\"m365_chili\"" >> ~/.bash_profile' \
  && bash -c 'echo ". .bashrc" >> ~/.bash_profile' \
  && bash -c 'npm cache clean --force' \
  && bash -c 'm365 cli completion sh setup' \
  && pwsh -c 'm365 cli completion pwsh setup --profile $profile' \
  && pwsh -c 'Add-Content -Path $PROFILE -Value "`r`Set-Alias -Name m365? -Value m365_chili"'

RUN pip install setuptools==58
RUN pip install jmespath-terminal

CMD [ "bash", "-l" ]