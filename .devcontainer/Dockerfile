FROM mcr.microsoft.com/powershell:7.1.0-ubuntu-focal

LABEL name="CLI for Microsoft 365 Development" \
  description="Development container for contributing to CLI for Microsoft 365" \
  homepage="https://pnp.github.io/cli-microsoft365" \
  maintainers="Waldek Mastykarz <waldek@mastykarz.nl>, \
  Velin Georgiev <velin.georgiev@gmail.com>, \
  Garry Trinder <garry.trinder@live.com>, \
  Albert-Jan Schot <appie@digiwijs.nl>, \
  Rabia Williams <rabiawilliams@gmail.com>, \
  Mark Powney <mpowney@icloud.com>, \
  Patrick Lamber <patrick@nubo.eu>"

ENV NPM_CONFIG_PREFIX=/home/cli-microsoft365-dev/.npm-global \
  PATH=$PATH:/home/cli-microsoft365-dev/.npm-global/bin

RUN apt-get update && apt-get install -y \
  git \
  curl \
  sudo \
  zsh \
  jq \
  vim \
  python3.8 \
  python3-pip \
  && curl -sL https://deb.nodesource.com/setup_14.x | sudo -E bash - \
  && apt-get install nodejs -y \
  && rm -rf /var/lib/apt/lists/*

RUN pip3 install mkdocs-material==3.1.0

RUN useradd \
  --user-group \
  --system \
  --create-home \
  --no-log-init \
  cli-microsoft365-dev
  
USER cli-microsoft365-dev

WORKDIR /home/cli-microsoft365

RUN zsh -c "$(curl https://raw.githubusercontent.com/robbyrussell/oh-my-zsh/master/tools/install.sh)" --unattended \
 && git clone https://github.com/denysdovhan/spaceship-prompt.git "$HOME/.oh-my-zsh/custom/themes/spaceship-prompt" --depth=1 \
 && ln -s "$HOME/.oh-my-zsh/custom/themes/spaceship-prompt/spaceship.zsh-theme" "$HOME/.oh-my-zsh/custom/themes/spaceship.zsh-theme" \
 && cp "$HOME/.oh-my-zsh/templates/zshrc.zsh-template" $HOME/.zshrc \
 && git clone https://github.com/zsh-users/zsh-syntax-highlighting.git $HOME/.oh-my-zsh/custom/plugins/zsh-syntax-highlighting \
 && git clone https://github.com/zsh-users/zsh-autosuggestions $HOME/.oh-my-zsh/custom/plugins/zsh-autosuggestions \
 && sed -i "11s/ZSH_THEME=\"robbyrussell\"/ZSH_THEME=\"spaceship\"/" ~/.zshrc \
 && sed -i "27s/# DISABLE_AUTO_UPDATE=\"true\"/DISABLE_AUTO_UPDATE=\"true\"/" ~/.zshrc \
 && sed -i "73s/plugins=(git)/plugins=(git zsh-autosuggestions zsh-syntax-highlighting)/" ~/.zshrc \
 && sed -i "50s/# COMPLETION_WAITING_DOTS/COMPLETION_WAITING_DOTS/" ~/.zshrc

CMD [ "/bin/zsh", "-l" ]