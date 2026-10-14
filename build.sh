#!/bin/bash
set -e

REGISTRY="cguevara962"
TRAIN_IMAGE="$REGISTRY/consumption-train:latest"
SERVE_IMAGE="$REGISTRY/consumption-serve:latest"

echo "==> Login DockerHub"
buildah login -u "$REGISTRY" --password "$DOCKERHUB_TOKEN" docker.io

echo "==> Build train"
buildah bud -f Dockerfile.train -t "$TRAIN_IMAGE" .

echo "==> Push train"
buildah push "$TRAIN_IMAGE"

echo "==> Build serve"
buildah bud -f Dockerfile.serve -t "$SERVE_IMAGE" .

echo "==> Push serve"
buildah push "$SERVE_IMAGE"

echo "==> Cleanup"
buildah rmi --all

echo "Done."
