FROM holgerbrandl/kscript:4.1.1

COPY src /usr/local/src

ENV KOTLIN_HOME=/opt/kotlinc
ENV PATH=/opt/kotlinc/bin:${PATH}

# 依存関係ダウンロード用に一度実行する
RUN /usr/bin/env kotlin /usr/local/src/converter.main.kts || true

WORKDIR /workspaces

ENTRYPOINT ["/usr/bin/env", "kotlin", "/usr/local/src/converter.main.kts"]
