"""Unit test suite for `pptx.media` module."""

from __future__ import annotations

import io

import pytest

from pptx.media import Audio, Video
from pptx.opc.constants import CONTENT_TYPE as CT

from .unitutil.file import absjoin, test_file_dir
from .unitutil.mock import initializer_mock, instance_mock, method_mock, property_mock

TEST_VIDEO_PATH = absjoin(test_file_dir, "dummy.mp4")


class DescribeVideo:
    """Unit-test suite for `pptx.media.Video` objects."""

    def it_can_construct_from_a_path(self, video_, from_blob_):
        with open(TEST_VIDEO_PATH, "rb") as f:
            blob = f.read()
        from_blob_.return_value = video_

        video = Video.from_path_or_file_like(TEST_VIDEO_PATH, "video/mp4")

        Video.from_blob.assert_called_once_with(blob, "video/mp4", "dummy.mp4")
        assert video is video_

    def it_can_construct_from_a_stream(self, from_stream_fixture):
        movie_stream, mime_type, blob, video_ = from_stream_fixture
        video = Video.from_path_or_file_like(movie_stream, mime_type)
        Video.from_blob.assert_called_once_with(blob, mime_type, None)
        assert video is video_

    def it_can_construct_from_a_blob(self, from_blob_fixture):
        blob, mime_type, filename, Video_init_ = from_blob_fixture
        video = Video.from_blob(blob, mime_type, filename)
        Video_init_.assert_called_once_with(video, blob, mime_type, filename)
        assert isinstance(video, Video)

    def it_provides_access_to_the_video_bytestream(self, blob_fixture):
        video, expected_value = blob_fixture
        assert video.blob == expected_value

    def it_knows_its_content_type(self, content_type_fixture):
        video, expected_value = content_type_fixture
        assert video.content_type == expected_value

    def it_knows_a_filename_for_the_video(self, filename_fixture):
        video, expected_value = filename_fixture
        assert video.filename == expected_value

    def it_knows_an_extension_for_the_video(self, ext_fixture):
        video, expected_value = ext_fixture
        assert video.ext == expected_value

    def it_knows_its_sha1_hash(self, sha1_fixture):
        video, expected_value = sha1_fixture
        assert video.sha1 == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def blob_fixture(self):
        blob = b"blob-bytes"
        video = Video(blob, None, None)
        expected_value = blob
        return video, expected_value

    @pytest.fixture
    def content_type_fixture(self):
        mime_type = "video/mp4"
        video = Video(None, mime_type, None)
        expected_value = mime_type
        return video, expected_value

    @pytest.fixture(
        params=[
            (None, "foo.bar", "bar"),
            ("video/mp4", None, "mp4"),
            ("video/xyz", None, "vid"),
        ]
    )
    def ext_fixture(self, request):
        mime_type, filename, expected_value = request.param
        video = Video(None, mime_type, filename)
        return video, expected_value

    @pytest.fixture(params=[("foobar.mp4", None, "foobar.mp4"), (None, "vid", "movie.vid")])
    def filename_fixture(self, request, ext_prop_):
        filename, ext, expected_value = request.param
        video = Video(None, None, filename)
        ext_prop_.return_value = ext
        return video, expected_value

    @pytest.fixture
    def from_blob_fixture(self, Video_init_):
        blob, mime_type, filename = "01234", "video/mp4", "movie.mp4"
        return blob, mime_type, filename, Video_init_

    @pytest.fixture
    def from_stream_fixture(self, video_, from_blob_):
        with open(TEST_VIDEO_PATH, "rb") as f:
            blob = f.read()
            movie_stream = io.BytesIO(blob)
        mime_type = "video/mp4"
        from_blob_.return_value = video_
        return movie_stream, mime_type, blob, video_

    @pytest.fixture
    def sha1_fixture(self):
        blob = b"blobish"
        video = Video(blob, None, None)
        expected_value = "de731a6eed12f427642325193b8e57af3c624d62"
        return video, expected_value

    # fixture components ---------------------------------------------

    @pytest.fixture
    def ext_prop_(self, request):
        return property_mock(request, Video, "ext")

    @pytest.fixture
    def from_blob_(self, request):
        return method_mock(request, Video, "from_blob", autospec=False)

    @pytest.fixture
    def video_(self, request):
        return instance_mock(request, Video)

    @pytest.fixture
    def Video_init_(self, request):
        return initializer_mock(request, Video, autospec=True)


class DescribeAudio:
    """Unit-test suite for `pptx.media.Audio` objects."""

    def it_can_construct_from_a_path(self, audio_, from_blob_, tmp_path):
        blob = b"01234-mp3-bytes"
        path = tmp_path / "clip.mp3"
        path.write_bytes(blob)
        from_blob_.return_value = audio_

        audio = Audio.from_path_or_file_like(str(path), "audio/mpeg")

        Audio.from_blob.assert_called_once_with(blob, "audio/mpeg", "clip.mp3")
        assert audio is audio_

    def it_can_construct_from_a_stream(self, from_stream_fixture):
        audio_stream, mime_type, blob, audio_ = from_stream_fixture
        audio = Audio.from_path_or_file_like(audio_stream, mime_type)
        Audio.from_blob.assert_called_once_with(blob, mime_type, None)
        assert audio is audio_

    def it_infers_mime_type_from_extension_when_not_specified(self, ext_infer_fixture):
        path, expected_mime_type = ext_infer_fixture
        audio = Audio.from_path_or_file_like(path, None)
        assert audio.content_type == expected_mime_type

    def it_falls_back_to_mp3_for_a_stream_with_no_mime_type(self):
        audio = Audio.from_path_or_file_like(io.BytesIO(b"abc"), None)
        assert audio.content_type == CT.MP3

    def it_can_construct_from_a_blob(self, from_blob_fixture):
        blob, mime_type, filename, Audio_init_ = from_blob_fixture
        audio = Audio.from_blob(blob, mime_type, filename)
        Audio_init_.assert_called_once_with(audio, blob, mime_type, filename)
        assert isinstance(audio, Audio)

    def it_provides_access_to_the_audio_bytestream(self, blob_fixture):
        audio, expected_value = blob_fixture
        assert audio.blob == expected_value

    def it_knows_its_content_type(self, content_type_fixture):
        audio, expected_value = content_type_fixture
        assert audio.content_type == expected_value

    def it_knows_a_filename_for_the_audio(self, filename_fixture):
        audio, expected_value = filename_fixture
        assert audio.filename == expected_value

    def it_knows_an_extension_for_the_audio(self, ext_fixture):
        audio, expected_value = ext_fixture
        assert audio.ext == expected_value

    def it_knows_its_sha1_hash(self, sha1_fixture):
        audio, expected_value = sha1_fixture
        assert audio.sha1 == expected_value

    # fixtures -------------------------------------------------------

    @pytest.fixture
    def blob_fixture(self):
        blob = b"blob-bytes"
        audio = Audio(blob, None, None)
        expected_value = blob
        return audio, expected_value

    @pytest.fixture
    def content_type_fixture(self):
        mime_type = "audio/mpeg"
        audio = Audio(None, mime_type, None)
        expected_value = mime_type
        return audio, expected_value

    @pytest.fixture(
        params=[
            (None, "foo.bar", "bar"),
            ("audio/mpeg", None, "mp3"),
            ("audio/wav", None, "wav"),
            ("audio/mp4", None, "m4a"),
            ("audio/x-m4a", None, "m4a"),
            ("audio/xyz", None, "aud"),
        ]
    )
    def ext_fixture(self, request):
        mime_type, filename, expected_value = request.param
        audio = Audio(None, mime_type, filename)
        return audio, expected_value

    @pytest.fixture(params=[("foobar.mp3", None, "foobar.mp3"), (None, "wav", "audio.wav")])
    def filename_fixture(self, request, ext_prop_):
        filename, ext, expected_value = request.param
        audio = Audio(None, None, filename)
        ext_prop_.return_value = ext
        return audio, expected_value

    @pytest.fixture(
        params=[
            ("foo.mp3", CT.MP3),
            ("foo.wav", CT.WAV),
            ("foo.m4a", CT.M4A),
            ("foo.xyz", CT.MP3),
            ("foo", CT.MP3),
        ]
    )
    def ext_infer_fixture(self, request, tmp_path):
        filename, expected_mime_type = request.param
        path = tmp_path / filename
        path.write_bytes(b"blob")
        return str(path), expected_mime_type

    @pytest.fixture
    def from_blob_fixture(self, Audio_init_):
        blob, mime_type, filename = "01234", "audio/mpeg", "audio.mp3"
        return blob, mime_type, filename, Audio_init_

    @pytest.fixture
    def from_stream_fixture(self, audio_, from_blob_):
        blob = b"01234-mp3-bytes"
        audio_stream = io.BytesIO(blob)
        mime_type = "audio/mpeg"
        from_blob_.return_value = audio_
        return audio_stream, mime_type, blob, audio_

    @pytest.fixture
    def sha1_fixture(self):
        blob = b"blobish"
        audio = Audio(blob, None, None)
        expected_value = "de731a6eed12f427642325193b8e57af3c624d62"
        return audio, expected_value

    # fixture components ---------------------------------------------

    @pytest.fixture
    def ext_prop_(self, request):
        return property_mock(request, Audio, "ext")

    @pytest.fixture
    def from_blob_(self, request):
        return method_mock(request, Audio, "from_blob", autospec=False)

    @pytest.fixture
    def audio_(self, request):
        return instance_mock(request, Audio)

    @pytest.fixture
    def Audio_init_(self, request):
        return initializer_mock(request, Audio, autospec=True)
