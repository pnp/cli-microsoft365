FROM mcr.microsoft.com/powershell:alpine-3.17

LABEL name="CLI for Microsoft 365 Development" \
  description="Development container for contributing to CLI for Microsoft 365" \
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
  Milan Holemans <Milan.Holemans@vanroey.be>"

ENV NPM_CONFIG_PREFIX=/home/cli-microsoft365-dev/.npm-global \
  PATH=$PATH:/home/cli-microsoft365-dev/.npm-global/bin

RUN apk add --no-cache \
  git \
  curl \
  sudo \
  shadow \
  zsh \
  jq \
  nodejs \
  npm

RUN useradd \
  --user-group \
  --system \
  --create-home \
  --no-log-init \
  cli-microsoft365-dev
  
USER cli-microsoft365-dev

WORKDIR /home/cli-microsoft365

RUN git clone https://github.com/ohmyzsh/ohmyzsh.git ~/.oh-my-zsh \
 && cp ~/.oh-my-zsh/templates/zshrc.zsh-template ~/.zshrc \
 && git clone https://github.com/denysdovhan/spaceship-prompt.git ${ZSH_CUSTOM:-~/.oh-my-zsh/custom}/themes/spaceship-prompt --depth=1 \
 && ln -s ${ZSH_CUSTOM:-~/.oh-my-zsh/custom}/themes/spaceship-prompt/spaceship.zsh-theme ${ZSH_CUSTOM:-~/.oh-my-zsh/custom}/themes/spaceship.zsh-theme \
 && sed -i "11s/ZSH_THEME=\"robbyrussell\"/ZSH_THEME=\"spaceship\"/" ~/.zshrc \
 && git clone https://github.com/zsh-users/zsh-syntax-highlighting.git ${ZSH_CUSTOM:-~/.oh-my-zsh/custom}/plugins/zsh-syntax-highlighting \
 && git clone https://github.com/zsh-users/zsh-autosuggestions ${ZSH_CUSTOM:-~/.oh-my-zsh/custom}/plugins/zsh-autosuggestions \
 && sed -i "73s/plugins=(git)/plugins=(git zsh-autosuggestions zsh-syntax-highlighting)/" ~/.zshrc \
 && sed -i "27s/# DISABLE_AUTO_UPDATE=\"true\"/DISABLE_AUTO_UPDATE=\"true\"/" ~/.zshrc \
 && sed -i "50s/# COMPLETION_WAITING_DOTS/COMPLETION_WAITING_DOTS/" ~/.zshrc

CMD [ "/bin/zsh", "-l" ]
